Attribute VB_Name = "change_picture_moude"
Option Explicit
Global addition_condition(7) As Integer
Global last_addition_condition As Integer
Global addition_condition_statue As Boolean
'附加的作图限制条件的输入语句号
Global temp_key  As Integer  '记录键值
Global temp_last_point(7)  As Integer
Global put_name_point%
Type temp_element
no As Integer '临时元的序号
new_or_old As Boolean '是否新生元，为判断是否能作撤消最后一步操作用
End Type
Global temp_point(15) As temp_element
Global temp_line(7) As Integer
Global temp_circle(7) As Integer
'Globalmouse_move_coord As POINTAPI
Global mouse_type%, open_move% '0 down,1 move,2 up
Global temp_string As String
Global inter_point_type As Boolean
Global temp_four_point_fig_type As Byte
Global display_temp_four_point_fig As Byte
Global temp_item0() As item0_data_type
Global last_temp_item0 As Integer
Type change_picture_start_no_data
 change_point_start_no As Integer
 current_point_no As Integer
 change_line_start_no As Integer
 current_line_no As Integer
 change_circle_start_no As Integer
 current_circle_no As Integer
 is_picture_change As Boolean
End Type
Global change_picture_start_no As change_picture_start_no_data
Type four_point_fig_type
      poi(3) As Integer
      p(3) As POINTAPI
End Type
Type change_point_type
c_p As Integer
run_time As Integer
End Type
Dim change_poi(26) As change_point_type
Global last_change_point As Integer
Global temp_four_point_fig As four_point_fig_type

Public Sub change_circle(move_coord As POINTAPI, similar_ratio!, d As Boolean)
Dim i%
Dim X&, Y&
If is_first_move Then
 '*** Call redraw_red_circle(Circle_for_change.c)
Else ' If
 '*** Draw_form.Circle (Circle_for_change.move.X, Circle_for_change.move.Y), _
     circ(Circle_for_change.c).data(0).radii, QBColor(12)
End If
Circle_for_change.move = add_POINTAPI(Circle_for_change.move, mouse_move_coord)

  Draw_form.Circle (Circle_for_change.move.X, Circle_for_change.move.Y), _
     Circle_for_change.radii, QBColor(12)
End Sub
Public Sub change_polygon(move_coord As POINTAPI, similar_ratio!, direction As Boolean)
Dim i%
Dim X&, Y&
If is_first_move Then
 For i% = 1 To Polygon_for_change.p(0).total_v - 1
 Call C_display_picture.draw_red_point(Polygon_for_change.p(0).v(i%))
 Call redraw_red_line(search_for_line_number_from_two_point(Polygon_for_change.p(0).v(i%), _
     Polygon_for_change.p(0).v(i% - 1), 0, 0))
 Next i%
 Call C_display_picture.draw_red_point(Polygon_for_change.p(0).v(0))
 If Polygon_for_change.p(0).total_v > 2 Then
 Call redraw_red_line(search_for_line_number_from_two_point(Polygon_for_change.p(0).v(0), _
     Polygon_for_change.p(0).v(Polygon_for_change.p(0).total_v - 1), 0, 0))
     is_first_move = False
 End If

Else ' If
Call draw_polygon(Polygon_for_change.p(0), 1)
End If
Polygon_for_change.similar_ratio = similar_ratio!
Polygon_for_change.direction = direction
Polygon_for_change.move.X = Polygon_for_change.move.X + mo_x&
Polygon_for_change.move.Y = Polygon_for_change.move.Y + mo_y&
Polygon_for_change.rote_angle = Polygon_for_change.rote_angle + A
If Polygon_for_change.rote_angle > 2 * PI Then
 Polygon_for_change.rote_angle = Polygon_for_change.rote_angle - 2 * PI
ElseIf Polygon_for_change.rote_angle < -2 * PI Then
 Polygon_for_change.rote_angle = Polygon_for_change.rote_angle + 2 * PI
End If
Polygon_for_change.p(0).coord_center.X = _
   Polygon_for_change.p(0).center.X + _
     Polygon_for_change.move.X
Polygon_for_change.p(0).coord_center.Y = _
  Polygon_for_change.p(0).center.Y + _
    Polygon_for_change.move.Y
For i% = 0 To Polygon_for_change.p(0).total_v - 1
If Polygon_for_change.direction = False Then
X& = Polygon_for_change.p(0).coord_center.X + _
     ((m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.X - _
       Polygon_for_change.p(0).center.X) * _
       Cos(Polygon_for_change.rote_angle) - _
        (m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.Y - _
         Polygon_for_change.p(0).center.Y) * _
          Sin(Polygon_for_change.rote_angle)) * _
           Polygon_for_change.similar_ratio
Y& = Polygon_for_change.p(0).coord_center.Y + _
     ((m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.X - _
        Polygon_for_change.p(0).center.X) * _
      Sin(Polygon_for_change.rote_angle) + _
        (m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.Y - _
          Polygon_for_change.p(0).center.Y) * _
           Cos(Polygon_for_change.rote_angle)) * _
          Polygon_for_change.similar_ratio
Else
X& = Polygon_for_change.p(0).coord_center.X - _
     ((m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.X - _
       Polygon_for_change.p(0).center.X) * _
       Cos(Polygon_for_change.rote_angle) + _
        (m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.Y - _
          Polygon_for_change.p(0).center.Y) * _
          Sin(Polygon_for_change.rote_angle)) * _
           Polygon_for_change.similar_ratio
Y& = Polygon_for_change.p(0).coord_center.Y + _
     ((m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.X - _
       Polygon_for_change.p(0).center.X) * _
      Sin(Polygon_for_change.rote_angle) + _
        (m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.Y - _
          Polygon_for_change.p(0).center.Y) * _
         Cos(Polygon_for_change.rote_angle)) * _
          Polygon_for_change.similar_ratio
End If
          Polygon_for_change.p(0).coord(i%).X = X&
          Polygon_for_change.p(0).coord(i%).Y = Y&
Next i%
Call draw_polygon(Polygon_for_change.p(0), 1)
End Sub
Public Sub change_polygon1(mo_x&, _
 mo_y&)
Dim i%
Dim X&, Y&
Dim t!, r&
Call draw_polygon(Polygon_for_change.p(0), 1)
Polygon_for_change.move.X = Polygon_for_change.move.X + mo_x&
Polygon_for_change.move.Y = Polygon_for_change.move.Y + mo_y&
r& = (line_for_move.coord(1).X - line_for_move.coord(0).X) ^ 2 + _
   (line_for_move.coord(1).Y - line_for_move.coord(0).Y) ^ 2
For i% = 0 To Polygon_for_change.p(0).total_v - 1
 t! = ((line_for_move.coord(0).X - _
        Polygon_for_change.p(0).coord(i%).X) * _
     (line_for_move.coord(0).X - line_for_move.coord(1).X) + _
       (line_for_move.coord(0).Y - _
         Polygon_for_change.p(0).coord(i%).Y) * _
         (line_for_move.coord(0).Y - line_for_move.coord(1).Y)) / r&
 X& = 2 * (line_for_move.coord(0).X + _
     t! * (line_for_move.coord(1).X - line_for_move.coord(0).X)) - _
       Polygon_for_change.p(0).coord(i%).X
 Y& = 2 * (line_for_move.coord(0).Y + _
      t! * (line_for_move.coord(1).Y - line_for_move.coord(0).Y)) - _
            Polygon_for_change.p(0).coord(i%).Y

          Polygon_for_change.p(0).coord(i%).X = X&
          Polygon_for_change.p(0).coord(i%).Y = Y&
Next i%
Call draw_polygon(Polygon_for_change.p(0), 1)

End Sub
Public Sub change_polygon2() '中心对称
Dim i%
Dim X&, Y&
Dim t!, r&
Call draw_polygon(Polygon_for_change.p(0), 1)
For i% = 0 To Polygon_for_change.p(0).total_v - 1
 X& = 2 * center_p.X - Polygon_for_change.p(0).coord(i%).X
 Y& = 2 * center_p.Y - Polygon_for_change.p(0).coord(i%).Y

          Polygon_for_change.p(0).coord(i%).X = X&
          Polygon_for_change.p(0).coord(i%).Y = Y&
Next i%
Call draw_polygon(Polygon_for_change.p(0), 1)

End Sub

Public Function set_polygon1(ByVal p1%, ByVal p2%, _
    n%, p As polygon, ByVal d As Boolean, no%, _
     ByVal no_reduce As Byte) As Byte
Dim i%, l%, j%
Dim A!
Dim v$
A! = PI * (n% - 2) / n%
v$ = Trim(str(180 * (n% - 2) / n%))
l% = search_for_line_number_from_two_point(p1%, p2%, 0, 0) ' condition, display)
p.total_v = n%
p.v(0) = p1%
 p.coord(0) = m_poi(p1%).data(0).data0.coordinate
p.v(1) = p2%
 p.coord(1) = m_poi(p2%).data(0).data0.coordinate
For i% = 2 To n% - 1
 MDIForm1.Toolbar1.Buttons(21).Image = 33
  ' ***Call init_Point0(last_conditions.last_cond(1).point_no)
  wenti_cond(no%).point_no(i%) = last_conditions.last_cond(1).point_no
  If d = False Then '旋转方向
    p.coord(i%).X = p.coord(i% - 1).X + _
    (p.coord(i% - 2).X - p.coord(i% - 1).X) * Cos(A!) - _
     (p.coord(i% - 2).Y - p.coord(i% - 1).Y) * Sin(A!)
   p.coord(i%).Y = p.coord(i% - 1).Y + _
    (p.coord(i% - 2).X - p.coord(i% - 1).X) * Sin(A!) + _
     (p.coord(i% - 2).Y - p.coord(i% - 1).Y) * Cos(A!)
  Else
    p.coord(i%).X = p.coord(i% - 1).X _
   (p.coord(i% - 2).X - p.coord(i% - 1).X) * Cos(A!) + _
     (p.coord(i% - 2).Y - p.coord(i% - 1).Y) * Sin(A!)
    p.coord(i%).Y = p.coord(i% - 1).Y - _
    (p.coord(i% - 2).X - p.coord(i% - 1).X) * Sin(A!) + _
      (p.coord(i% - 2).Y - p.coord(i% - 1).Y) * Cos(A!)
  End If
   p.v(i%) = set_point(p.coord(i%), 1, condition_color, 0, wenti_cond(no%).condition(i%), 0)
Next i%
   For i% = 1 To p.total_v - 1
   Call line_number(p.v(i%), p.v((i% + 1) Mod p.total_v), condition, display)
   Next i%
If p.total_v = 4 Then
For i% = 1 To last_conditions.last_cond(1).point_no
 If i% <> p.v(0) And i% <> p.v(1) Then
  If is_dverti0(line_number5(i%, p.v(0), 0, 0, 0), l%) Then
   Call add_point_to_line(p.v(3), _
      line_number(p.v(0), i%, condition, display), 0, _
        display, True, 0, temp_record)
  ElseIf is_dverti0(line_number5(i%, p.v(1), 0, 0, 0), l%) Then
   Call add_point_to_line(p.v(2), _
      line_number(p.v(1), i%, condition, display), 0, _
       display, True, 0)
  End If
 End If
Next i%
End If

End Function


Public Function change_picture(ByVal num As Integer, change_element As condition_type, Optional ByVal change_ratio_for_measure&) As Boolean
'如果数一出false, num change_point,change_ratio_for_measure
'用于重画图
Dim o As Integer
Dim A!, b!, c!, d!, p!, q!, r!, t!, s!, k%, l%, n%, m%
Dim i%, j%, u%, v%, w!, X&, Y&, x1&, y1&
Dim e!, f!, g!, H!
Dim t_l%, t_l1%
Dim tn(3) As Long
Dim tp%, c_p%
Dim coord As POINTAPI
Dim coord1 As POINTAPI
Dim ty As Boolean
Dim wenti_data As wentitype
Call C_display_wenti.Get_wenti(num) '获取wenti_cond0即输入语句数据
wenti_data = wenti_cond0.data
Dim t_wenti_cond As wentitype
'poi(yidian_no).data(0).is_change
If num = 0 Then
   Exit Function
End If
If wenti_data.no_ = -1000 Then
    change_picture = True
     Exit Function
End If
Select Case wenti_data.no
Case -54, -22, -23
change_picture = change_picture_54_23_22(wenti_data, change_element)
'Case -51, -52, -56
'change_picture = change_picture_51_52_56(wenti_data, change_element)
'Case -51
'change_picture = change_picture_51(wenti_data)
Case -50, -51, -52, -56
change_picture = change_picture_50(wenti_data, change_element)
Case -43
 change_picture = change_picture_43(wenti_data)
Case -42, -57
 change_picture = change_picture_42_57(wenti_data)
Case -33, -44
 change_picture = change_picture_33_44(wenti_data)
Case -32, -27
 change_picture = change_picture_32(wenti_data)
'Case -31, -30, -58
' change_picture = change_picture_31_30(wenti_data)
Case -15
 change_picture = change_picture_15(wenti_data)
Case -14, -13, -11, -10
 change_picture = change_picture_14_13_11_10(wenti_data)
Case -3, -29, -28, -26
 change_picture = change_picture_3(wenti_data)
Case -18
 change_picture = change_picture_18(wenti_data)
Case -16, -12, -9, -8
 change_picture = change_picture_16_12_9_8(wenti_data, change_element)
Case -17
 change_picture = change_picture_17(wenti_data)
Case -6
 change_picture = change_picture_6(wenti_data, change_element)
Case -4
 change_picture = change_picture_4(wenti_data, change_element, c_p%)
Case -2, -60, -59, -33, -44
 change_picture = change_picture_2(wenti_data)
Case -1, -30, -31, -58
 change_picture = change_picture_1_30_31_58(wenti_data, change_element)
Case 0
 'change_picture = change_picture0(num)
Case 1 '线上任取一点
 change_picture = change_picture1(wenti_data, change_element, c_p%, num)
Case 2, 3 ' 平行上任取一点
 change_picture = change_picture2_3(wenti_data)
Case 4 ' 垂直平分'　中点应记录
 change_picture = change_picture4(wenti_data, change_element, c_p%)
Case 5 '中点
 change_picture = change_picture5(wenti_data, change_element, c_p%)
Case 6 '　定比分点  ?????
 change_picture = change_picture6(wenti_data)
'If c_display_wenti.m_point_no(0) > yidian_no Then
'End If
Case 7, -61 '圆上任取一点
 change_picture = change_picture7_61(wenti_data)
'Case 8, -71
 'change_picture = change_picture8_71(wenti_data, num, change_element, c_p%)
Case 9 '相交
 change_picture = change_picture9(wenti_data, change_element, c_p%)
Case 10, 16, -53, -25 '过□垂直□□的直线交⊙□于□,□□垂直平分线交⊙□于□
 change_picture = change_picture10_16(wenti_data, change_element)
Case 11, -63
 change_picture = change_picture11(wenti_data)
    '直线与圆的一个交点
Case 12, -65, -64
'-65 ⊙□□□和⊙□□□相切于点□
'-64 ⊙□□□和⊙□[down\\(_)]相切于点□
 change_picture = change_picture12(wenti_data)
Case 13, -66, -67
 change_picture = change_picture13(wenti_data)
Case 14 '垂足
 change_picture = change_picture14(wenti_data)
Case 15
 change_picture = change_picture15(wenti_data)
'inpcond(16) = "□等于是一元二次方程_的两根和"
'inpcond(17) = "□等于是一元二次方程_的两根积"
Case 18     '□是△□□□的重心
 change_picture = change_picture18(wenti_data)
Case 19 '外心
 change_picture = change_picture19(wenti_data)
Case 20 '垂心
 change_picture = change_picture20(wenti_data)
Case 21 '内心
 change_picture = change_picture21(wenti_data)
Case 23
 change_picture = change_picture23(wenti_data)
Case 29
 change_picture = change_picture29(wenti_data)
End Select
''Call change_picture_0(c_p%, wenti_data.point_no(29), num)
End Function
Sub draw_again1(ob As Object) ', display_or_delete As Boolean)
Dim i%, j%
'For i% = 1 To last_conditions.last_cond(1).point_no '***
'    Call set_point_color(i%, 0)
' Call draw_point(Ob, poi(i%), 0, display)
'Next i%
'For i% = 1 To last_conditions.last_cond(1).line_no  '***
'Call draw_line(Ob, m_lin(i%).data(0).data0, condition, 0)   ', condition)
'Next i%
'For i% = 1 To last_conditions.last_cond(1).con_line_no  '***
'Call draw_line(Ob, Con_lin(i%).data(0).data0, concl, 0)
'Next i%
'For i% = 1 To last_conditions.last_cond(1).circle_no '***
'Call draw_circle(ob, m_circ(i%).data(0).data0)
'Next i%
End Sub
Sub draw_again0(ob As Object, ty As Byte) 'ty=1 显示轨迹
Dim i%
For i% = 1 To last_conditions.last_cond(1).line_no  '***
If m_lin(i%).data(0).data0.visible = 1 And m_lin(i%).data(0).is_change = 255 Then
     m_lin(i%).data(0).data0.type = condition
        Call C_display_picture.set_m_line_data0(i%, 0, 0)
         'm_lin(i%).data(0).is_change = False
End If
Next i%
'For i% = 1 To last_conditions.last_cond(1).con_line_no  '***
'Call simple_con_line(m_Con_lin(i%).data(0).data0)
'Next i%
For i% = 1 To last_conditions.last_cond(1).point_no '***
If m_poi(i%).data(0).is_change Then
     'm_point_data0 = m_poi(i%).data(0).data0
       Call C_display_picture.PointCoordinateChange(i%)
        m_poi(i%).data(0).is_change = False
End If
Next i%
For i% = 1 To last_conditions.last_cond(1).circle_no '***
    If m_Circ(i%).data(0).is_change Then
      m_input_circle_data0 = m_Circ(i%).data(0)
        'Call C_display_picture.set_m_circle_data0(i%, 0)
          m_Circ(i%).data(0).is_change = False
    End If
Next i%
End Sub

Sub draw_picture(num As Integer, ByVal no_reduce As Byte, ByVal input_type As Boolean)
Dim i% '由输入作图
For i% = 0 To 50
  If Asc(C_display_wenti.m_condition(num, i%)) > 64 And _
           Asc(C_display_wenti.m_condition(num, i%)) < 91 Then
      Call C_display_wenti.set_m_point_no(num, _
         point_number(C_display_wenti.m_condition(num, i%)), i%, True)
         '读出条件的点, 记录点号
   End If
Next i%
'End If
If is_old_conclusion(num) Then
   Exit Sub
End If
Dim wenti_data As wentitype
Call C_display_wenti.Get_wenti(num) '获取wenti_cond0即输入语句数据
wenti_data = wenti_cond0.data
'*******************************************************
Select Case wenti_data.no 'C_display_wenti.m_no(num)
Case -52
Call draw_picture_52(num, no_reduce)
Case -51
Call draw_picture_51(num, no_reduce)
Case -50
Call draw_picture_50(num, no_reduce)
Case -48, -47
Call draw_picture_47_48(num, no_reduce)
Case -46, -45
Call draw_picture_45_46(num, no_reduce)
'Case -43, -42
'Call draw_picture_43_42(num, wenti_data, no_reduce)
Case -41 '∠□□□/∠□□□=!_~
Call draw_picture_41(num, no_reduce)
Case -40
Call draw_picture_40(num, no_reduce)
'□□/□□=□□/□□
Case -39
'inpcond(-39) = ∠□□□=∠□□□+∠□□□
Call draw_picture_39(num, no_reduce)
Case -38 '∠□□□+∠□□□=!_~°
Call draw_picture_38(num, no_reduce)
Case -37 '△□□□≌△□□□
Call draw_picture_37(num, no_reduce)
Case -36 '△□□□∽△□□□
Call draw_picture_36(num, no_reduce)
Case -35 '□□=□□+□□
Call draw_picture_35(num, no_reduce)
Case -34 '□□+□□=!_~
Call draw_picture_34(num, no_reduce)
Case -33
'inpcond(-33) = loadresstring_(304) '□□⊥□□
Call draw_picture_33(num, no_reduce)
Case -32, -3
'inpcond(-32) = loadresstring_(305) '□□∥□□
Call draw_picture_32_3(num, no_reduce)
Case -31
'inpcond(-31) = loadresstring_(306)
Call draw_picture_31(num, no_reduce)
Case -30
'inpcond(-30) = loadresstring_(307)
Call draw_picture_30(num, no_reduce)
Case -24
'inpcond(-24) = 弧□□=弧□□
Call draw_picture_24(num, no_reduce)
Case -23, -22
'inpcond(-23) = 过□点垂直□□的直线交□□于□
'inpcond(-22) = 过□点平行□□的直线交□□于□
Call draw_picture_23_22(num, no_reduce)
Case -20
'inpcond(-21) =△□□□是直角三角形
'inpcond(-20) = 任意△□□□
Call draw_picture_20(num, no_reduce)
Case -21
Call draw_picture_21(num, no_reduce)
Case -19
Call draw_picture_19(num, no_reduce)
Case -17, -18
If draw_picture_17_18(num, no_reduce) Then
 Exit Sub
End If
Case -16, -12, -9, -8
Call draw_picture_16_12_9_8(num, no_reduce)
'*******************************************************
Case -15
Call draw_picture_15(num, no_reduce)
'******************************************************************
Case -13, -11, -14, -10
Call draw_picture_13_11_14_10(num, no_reduce)
Case -7
Call draw_picture_7(num, no_reduce)
Case -6, -43, -42, -57
Call draw_picture_6_43_42_57(num, no_reduce)
Case -5
Call draw_picture_5(num, no_reduce)
Case -4
Call draw_picture_4(num, no_reduce)
Case -2
 Call draw_picture_2(num, no_reduce)
Case -1
 Call draw_picture_1(num, 0)
Case 0
 Call draw_picture0(num, no_reduce)
Case 1 '线上任取一点
Call draw_picture1(num, no_reduce)
Case 2, 3  ' 平行垂直上任取一点'***
Call draw_picture2_3(num, no_reduce)
Case 4 ' 垂直平分
'新点应加在最后
Call draw_picture4(num, no_reduce)
Case 5, 15 '中点
Call draw_picture5_15(num, no_reduce)
Case 6 '定比分点
Call draw_picture6(num, no_reduce)
Case 7 ' 圆上任取一点
Call draw_picture7(num, no_reduce)
Case 8
Call draw_picture8(num, no_reduce)
Case 9 '两直线相交
Call draw_picture9(num, no_reduce)
Case 10, 16 '直线与圆已交于一点求另一交点
Call draw_picture10_16(num, no_reduce)
Case 11, -63
Call draw_picture11(num, no_reduce)
Case 13  '两圆的一个交点
Call draw_picture13(num, no_reduce)
Case 14  '过□作直线□□的垂线垂足为□
Call draw_picture14(num, no_reduce)
'Case 17, 15, 16 '平行   垂直
'Call draw_picture15_16_17(num, no_reduce)
Case 18, 19, 20, 21 '□是△□□□的重心
If draw_picture18_21(num, no_reduce) Then
 Exit Sub
End If
'Case 22
'Call draw_picture22(num, no_reduce)
Case 23  '□、□、□、□四点共圆
Call draw_picture23(num)
Case 24  '□、□、□三点共线
Call draw_picture24(num)
Case 25, 27, 28
Call draw_picture25_27_28(num)
'"线段□□和□□长相等，即｜□□｜＝｜□□｜"
'"□□平行于□□"
' "□□垂直于□□"
'i% = line_number(c_display_wenti.m_point_no(0), _
 'c_display_wenti.m_point_no(1), concl, display)
'Call add_point_to_con_line(c_display_wenti.m_point_no(2), i%)
Case 29
Call draw_picture29(num)
Case 26
'点□是线段□□的中点
Call draw_picture26(num)
Case 31
'"线段□□上的分点□满足□□：□□＝_"
Call draw_picture31(num)
Case 30
Call draw_picture30(num)
'∠□□□=∠□□□
'inpcond(31) = □□:□□=!_~
'inpcond(32) = □□/□□=□□/□□
'inpcond(33) = △□□□∽△□□□
'inpcond(34) = △□□□≌△□□□
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Case 32
Call draw_picture32(num)
Case 33, 34
Call draw_picture33_34(num)
Case 35, 54
Call draw_picture35_54(num)
Case 36, 53
Call draw_picture36_53(num)
Case 37
Call draw_picture37(num)
Case 38, -49, 50
Call draw_picture38_49(num, 0, 19, no_reduce)
Case 39
Call draw_picture39(num)
Case 40
Call draw_picture40(num)
Case 41
Call draw_picture41(num)
Case 42
Call draw_picture42(num)
Case 43
Call draw_picture43(num)
Case 44
Call draw_picture44(num)
Case 45
Call draw_picture45(num)
Case 46
Call draw_picture46(num)
Case 47
Call draw_picture47(num)
Case 49
Call draw_picture49(num)
Case 51
Call draw_picture51(num)
Case 52
Call draw_picture52(num)
Case 55
Call draw_picture55(num)
Case 56
Call draw_picture56(num)
Case 57
Call draw_picture57(num)
Case 58
Call draw_picture58(num)
Case 59
Call draw_picture59(num)
Case 60, 62
Call draw_picture60_62(num)
Case 61
Call draw_picture61(num)
'Case 62
'Call draw_picture60_62(num)
Case 63
Call draw_picture63(num)
Case 64, 66
Call draw_picture64_66(num)
Case 65
Call draw_picture65(num)
Case 65
Call draw_picture65(num)
Case 68, 69
Call draw_picture68_69(num)
End Select
MDIForm1.Timer1.Enabled = False
draw_wenti_no = num + 1
'event_statue = ready
'If c_display_wenti.m_no < 23 Then
'Call call_theorem(0)
'End If
End Sub


Public Sub draw_picture58(ByVal num As Integer)
'58 求⊙□[down\\(_)]的面积
Dim A As Integer
Dim i%
Dim value1 As String
If is_old_conclusion(num) Then
   Exit Sub
End If
conclusion_data(last_conclusion).ty = area_of_circle_
A = m_circle_number(1, C_display_wenti.m_point_no(num, 0), pointapi0, _
       C_display_wenti.m_point_no(num, 1), 0, 0, 0, 0, 0, 1, 1, _
         conclusion, conclusion_color, True)
    con_Area_of_circle(last_conclusion).data(0).circ = A
      conclusion_data(last_conclusion).wenti_no = num
        last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True

End Sub

Sub point_position_on_circle(ByVal p%, ByVal c%, A%, b%)
Dim i%
Dim r!
A% = Int(CSng(m_poi(p%).data(0).data0.coordinate.X - m_Circ(c%).data(0).data0.c_coord.X) * 1000 _
      \ m_Circ(c%).data(0).data0.radii)
b% = Int(CSng(m_poi(p%).data(0).data0.coordinate.Y - m_Circ(c%).data(0).data0.c_coord.Y) * 1000 _
     \ m_Circ(c%).data(0).data0.radii)
End Sub

Sub put_name_to_point(ByVal p%, ByVal w%, ByVal n%)
'为点命名
Dim char As String * 1
'Call draw_point(p%)

'poi(p%).data(0).data0.color = 12
Draw_form.Label2.top = m_poi(p%).data(0).data0.coordinate.Y - 20
 Draw_form.Label2.left = m_poi(p%).data(0).data0.coordinate.X + 4

Draw_form.Label2.visible = True
Draw_form.Label2.Caption = LoadResString_(1725, "")

'显示
'time_no = 0

'Draw_form.Timer1.Enabled = 1

'Call draw_point(p%)
put_name_back:
input_statue_from_p = 1
While input_statue_from_p = 1
DoEvents
Wend

char = Chr(temp_key)
Call C_display_wenti.set_m_condition(w%, char, n%)
Call C_display_wenti.set_m_point_no(w%, p%, n%, False)
If get_input_info(w%, n%) Then
'Draw_form.timer1.Enabled = 1
Draw_form.Label2.Caption = input_char_info
GoTo put_name_back
End If
Draw_form.Label2.visible = False
Call set_point_name(p%, char)
Call C_display_picture.set_m_point_color(p%, 0)
'Call draw_point(Draw_form, poi(p%), 0, display)
put_name_point% = 0
End Sub

Sub put_point_number()
Dim i%, j%
For j% = 0 To 30
For i = 0 To 20
If Asc(C_display_wenti.m_condition(j%, i%)) > 63 And _
      Asc(C_display_wenti.m_condition(j%, i%)) < 91 Then         '  63 And Asc(c_display_wenti.m_condition(i)) < 91
Call C_display_wenti.set_m_point_no(j%, _
     point_number(C_display_wenti.m_condition(j%, i)), i%, True)
Else
Call C_display_wenti.set_m_point_no(j%, 0, i%, False)
End If
Next i%
Next j%
End Sub

Sub temp_init()
Dim i%

For i% = 0 To 7
temp_last_point(i%) = 0
Next i%
For i% = 0 To 15
temp_point(i%).no = 0
temp_point(i%).new_or_old = False
Next
For i% = 0 To 3
temp_circle(0) = 0
temp_line(i%) = 0
Next
End Sub


Public Sub simple_angle(ByVal i%)
Dim l%, n%, tn%
Dim A As angle_type
n% = Abs(angle_number(angle(i%).data(0).poi(0), angle(i%).data(0).poi(1), angle(i%).data(0).poi(2), 0, 0))
 If i% = n% Then
  Exit Sub
 Else
  ' angle(i%).data(0).other_no = n%
    angle(i%).data(0).no_reduce = 2
 End If
End Sub
Public Sub draw_picture_1(ByVal num%, ByVal no_reduce As Byte) ', p5%, p6%, p7%, p8%) ', ratio1%, ratio2%)
'-1 □□＝□□
Dim i%
Dim tp(3) As Integer
   For i% = 0 To 3
    If draw_free_point(C_display_wenti.m_point_no(num, i%), _
      C_display_wenti.m_condition(num, i%)) Then
       Exit Sub
    End If
   Next i%
 End Sub


Public Sub draw_p0(num%, i%) 'p%, ByVal n$)
'画点
Dim t_ele1 As condition_type
Dim t_ele2 As condition_type
Dim p%, t_point1%
Dim n$
p% = C_display_wenti.m_point_no(num, i%)
 n$ = C_display_wenti.m_condition(num, i%)
draw_p0_mark0:
  event_statue = wait_for_draw_point '输点状态
   While event_statue = wait_for_draw_point '等待事件发生
    DoEvents
   Wend
 If event_statue = draw_point_down Or event_statue = _
             draw_point_move Or event_statue = _
                    draw_point_up Then 'mouse_type <> 1 Then
   t_coord = input_coord
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_p0_mark0
      End If
      temp_point(0).no = 0
input_point_type% = read_inter_point(t_coord, t_ele1, t_ele2, temp_point(0).no, True)
          Call set_point_no_reduce(temp_point(0).no, 0)
     If input_point_type% <> new_free_point Then  '不是新的自由点
         If input_point_type% <> exist_point Then  '不是旧的自由点
          Call remove_point(temp_point(0).no, display, 0) '抹掉
         End If
          GoTo draw_p0_mark0
     End If
       Call set_point_name(temp_point(0).no, n$)
       Call C_display_wenti.set_m_point_no(num%, temp_point(0).no, i%, False)
         p% = temp_point(0).no
       ' Call put_name(p%)
End Sub
Public Sub draw_p3_4(p1%, p2%, p3%, ty As Boolean, p4%, name$, ratio!)
Dim tp%
Dim t_ele1 As condition_type
Dim t_ele2  As condition_type
Dim temp_record As total_record_type
    temp_line(0) = line_number(p2%, p3%, pointapi0, pointapi0, _
                     depend_condition(point_, p2%), depend_condition(point_, p3%), _
                      condition, condition_color, 1, 0)                                       '(k%, l%)
    last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1 '***
    MDIForm1.Toolbar1.Buttons(21).Image = 33
'     Call init_Point0(last_conditions.last_cond(1).point_no)
      temp_point(7).no = last_conditions.last_cond(1).point_no
     If ty Then
       t_coord = add_POINTAPI(m_poi(p1%).data(0).data0.coordinate, _
                     minus_POINTAPI(m_poi(p3%).data(0).data0.coordinate, _
                      m_poi(p2%).data(0).data0.coordinate))
         Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
     Else 'If c_display_wenti.m_no = 3 Then
       t_coord.X = m_poi(p1%).data(0).data0.coordinate.X + _
         m_poi(p3%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y
       t_coord.Y = m_poi(p1%).data(0).data0.coordinate.Y + _
         m_poi(p2%).data(0).data0.coordinate.X - m_poi(p3%).data(0).data0.coordinate.X
            Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
     End If
 '*******************************************************
        Call set_point_visible(last_conditions.last_cond(1).point_no, 0, False)
     If ty Then
      If is_point_in_line3(p1%, m_lin(temp_line(0)).data(0).data0, 0) Then
        temp_line(1) = temp_line(0)
        ' Call add_point_to_line(temp_point(2), temp_line(1), True, display)
      Else
       temp_line(1) = line_number(p1%, last_conditions.last_cond(1).point_no, _
                                  pointapi0, pointapi0, _
                                  depend_condition(point_, p1%), _
                                  depend_condition(point_, last_conditions.last_cond(1).point_no), _
                                  condition, condition_color, 1, 0)
      Call paral_line(temp_line(0), temp_line(1), True, True)  '
       End If
     Else 'If c_display_wenti.m_no = 3 Then
      temp_line(1) = line_number(p1%, last_conditions.last_cond(1).point_no, _
                                 pointapi0, pointapi0, _
                                 depend_condition(point_, p1%), _
                                 depend_condition(point_, last_conditions.last_cond(1).point_no), _
                                 condition, condition_color, 1, 0)
       Call vertical_line(temp_line(0), temp_line(1), True, True) '
      End If
draw_picture1_mark2:
    event_statue = wait_for_draw_point
     While event_statue = wait_for_draw_point
     DoEvents
      Wend
 If event_statue = draw_point_down Or _
    event_statue = draw_point_move Or _
        event_statue = draw_point_up Then 'mouse_type <> 1 Then
   t_coord = input_coord
    'temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark2
 End If
     input_point_type% = read_inter_point(t_coord, t_ele1, _
                                     t_ele2, temp_point(0).no, True)
         Call set_point_no_reduce(temp_point(0).no, 0)
         If input_point_type% <> new_point_on_line Or _
             t_ele1.no <> temp_line(1) Then
         If input_point_type% <> exist_point Then  '不是旧的自由点
           Call remove_point(temp_point(0).no, display, 0)
          GoTo draw_picture1_mark2
          Else 'End If
          GoTo draw_picture1_mark2
          End If
         End If
 'End If
  Call remove_point(temp_point(7).no, display, 0)
    Call set_point_name(temp_point(0).no, name)
      m_poi(temp_point(0).no).data(0).degree = 1
       Call set_point_in_line(temp_point(0).no, temp_line(1))
           p4% = temp_point(0).no
     If ty = False Then
          If Abs(m_poi(p3%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) > 4 Then
           ratio! = -(m_poi(p4%).data(0).data0.coordinate.Y - _
             m_poi(p1%).data(0).data0.coordinate.Y) / (m_poi(p3%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X)
          Else
           ratio! = (m_poi(p4%).data(0).data0.coordinate.X - _
              m_poi(p1%).data(0).data0.coordinate.X) / (m_poi(p3%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y)
          End If
     Else 'If c_display_wenti.m_no = 2 Then
          If Abs(m_poi(p3%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) > 4 Then
           ratio! = (m_poi(p4%).data(0).data0.coordinate.X - _
               m_poi(p1%).data(0).data0.coordinate.X) / (m_poi(p3%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X)
          Else
           ratio! = (m_poi(p4%).data(0).data0.coordinate.Y - _
               m_poi(p1%).data(0).data0.coordinate.Y) / (m_poi(p3%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y)
          End If
     End If
         ' c_display_wenti.m_point_no(4) = Int(A! * 1000)

'End If

End Sub




Public Sub set_wenti1(ByVal n%, p1%, p2%, p3%, l%)
Dim A%
 If C_display_wenti.m_condition(n%, 0) = "" Or _
      C_display_wenti.m_condition(n%, 0) = empty_char Then
     Call C_display_wenti.set_m_condition(n%, m_poi(p1%).data(0).data0.name, 0)
     Call C_display_wenti.set_m_point_no(n%, p1%, 0, True)
 End If
 If C_display_wenti.m_condition(n%, 1) = "" Or _
     C_display_wenti.m_condition(n%, 1) = empty_char Then
     Call C_display_wenti.set_m_condition(n%, m_poi(p2%).data(0).data0.name, 1)
     Call C_display_wenti.set_m_point_no(n%, p2%, 1, True)
 End If
  If C_display_wenti.m_condition(n%, 2) = "" Or _
     C_display_wenti.m_condition(n%, 2) = empty_char Then
     Call C_display_wenti.set_m_condition(n%, m_poi(p2%).data(0).data0.name, 2)
     Call C_display_wenti.set_m_point_no(n%, p3%, 2, True)
 End If
 If l% = 0 Then
  l% = line_number(p1%, p2%, pointapi0, pointapi0, _
                   depend_condition(point_, p1%), _
                   depend_condition(point_, p2%), _
                   condition, condition_color, 1, 0)
 End If
   If C_display_wenti.m_point_no(n%, 4) = 0 Then
      Call C_display_wenti.set_m_point_no(n%, l%, 4, True)
   End If
        'm_poi(p3%).data(0).in_line(0) = m_poi(p2%).data(0).in_line(0) + 1
         Call set_point_in_line(p3%, l%)
         m_poi(p3%).data(0).degree = 1
         record_0.data0.condition_data.condition_no = 0
         Call add_point_to_line(p3%, l%, 0, display, True, 0, temp_record)
     Call point_to_ratio(m_poi(p1%).data(0).data0.coordinate, _
        m_poi(p3%).data(0).data0.coordinate, m_poi(p2%).data(0).data0.coordinate, A%)
     Call C_display_wenti.set_m_point_no(n%, A%, 3, False)
     Call C_display_wenti.set_m_no(0, n%, 1)
End Sub


Public Sub draw_picture_33(ByVal num As Integer, ByVal no_reduce As Byte)
'If c_display_wenti.m_point_no(num,3) > 0 And c_display_wenti.m_point_no(num,4) Then
'   temp_line(2) = line_number(c_display_wenti.m_point_no(num,3), _
           c_display_wenti.m_point_no(num,4), condition, display, 0)
'   If c_display_wenti.m_point_no(num,11) = 0 Then
'      For i% = 1 To Circ(c_display_wenti.m_point_no(num,8)).data(0).data0.in_point(0)
'       For j% = 1 To lin(temp_line(2)).data(0).data0.in_point(0)
'           If Circ(c_display_wenti.m_point_no(num,8)).data(0).data0.in_point(i%) = _
'                lin(temp_line(2)).data(0).data0.in_point(j%) Then
'                 c_display_wenti.m_point_no(num,11) = lin(temp_line(2)).data(0).data0.in_point(j%)
'                  GoTo draw_picture_33_next1
'           End If
'       Next j%
'      Next i%
'   End If
'draw_picture_33_next1:
' MDIForm1.Toolbar1.Buttons(21).Image = 33
' If c_display_wenti.m_point_no(num,11) = lin(temp_line(2)).data(0).data0.poi(0) Then
' last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'   Call init_Point0(last_conditions.last_cond(1).point_no)
'    temp_point(0) = last_conditions.last_cond(1).point_no
'    poi(temp_point(0)).data(0).data0.coordinate.X = _
'      2 * poi(lin(temp_line(2)).data(0).data0.poi(0)).data(0).data0.coordinate.X - _
          poi(lin(temp_line(2)).data(0).data0.poi(1)).data(0).data0.coordinate.X
'    poi(temp_point(0)).data(0).data0.coordinate.Y = _
'      2 * poi(lin(temp_line(2)).data(0).data0.poi(0)).data(0).data0.coordinate.Y - _
'          poi(lin(temp_line(2)).data(0).data0.poi(1)).data(0).data0.coordinate.Y
'    Call put_name(temp_point(0))
'    Call add_point_to_line(temp_point(0), temp_line(2), 0, False, False)
' ElseIf c_display_wenti.m_point_no(num,11) = lin(temp_line(2)).data(0).data0.poi(1) Then
' last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'   Call init_Point0(last_conditions.last_cond(1).point_no)
'    temp_point(0) = last_conditions.last_cond(1).point_no
'    poi(temp_point(0)).data(0).data0.coordinate.X = _
'      2 * poi(lin(temp_line(2)).data(0).data0.poi(1)).data(0).data0.coordinate.X - _
'          poi(lin(temp_line(2)).data(0).data0.poi(0)).data(0).data0.coordinate.X
'    poi(temp_point(0)).data(0).data0.coordinate.Y = _
'      2 * poi(lin(temp_line(2)).data(0).data0.poi(1)).data(0).data0.coordinate.Y - _
          poi(lin(temp_line(2)).data(0).data0.poi(0)).data(0).data0.coordinate.Y
'    Call put_name(temp_point(0))
'    Call add_point_to_line(temp_point(0), temp_line(2), 0, False, False)
' End If
Dim i%, j%, k%
Dim ele1 As condition_type
Dim ele2  As condition_type
'Dim temp_x&, temp_y&
Dim A!
Call C_display_wenti.set_m_point_no(num, _
       m_circle_number(1, C_display_wenti.m_point_no(num, 1), _
        pointapi0, C_display_wenti.m_point_no(num, 2), _
         0, 0, 0, 0, 0, 1, 1, condition, condition_color, True), 8, False)
If is_point_in_circle(C_display_wenti.m_point_no(num, 8), _
                    0, C_display_wenti.m_point_no(num, 0), 0, 0) Then
   Call C_display_wenti.set_m_point_no(num, _
         C_display_wenti.m_point_no(num, 0), 11, False)
End If
If C_display_wenti.m_point_no(num, 3) > 0 And C_display_wenti.m_point_no(num, 0) > 0 Then
   temp_line(2) = line_number(C_display_wenti.m_point_no(num, 3), _
                              C_display_wenti.m_point_no(num, 0), _
                              pointapi0, pointapi0, _
                              depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                              depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                              condition, condition_color, 1, 0)
    If C_display_wenti.m_point_no(num, 11) = 0 Then
       Call C_display_wenti.set_m_point_no(num, _
        C_display_wenti.m_point_no(num, 3), 11, False)
    End If
 MDIForm1.Toolbar1.Buttons(21).Image = 33
 If C_display_wenti.m_point_no(num, 11) = m_lin(temp_line(2)).data(0).data0.poi(0) Then
  last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
  ' Call init_Point0(last_conditions.last_cond(1).point_no)
    temp_point(0).no = last_conditions.last_cond(1).point_no
      t_coord.X = 2 * m_poi(m_lin(temp_line(2)).data(0).data0.poi(0)).data(0).data0.coordinate.X - _
          m_poi(m_lin(temp_line(2)).data(0).data0.poi(1)).data(0).data0.coordinate.X
      t_coord.Y = 2 * m_poi(m_lin(temp_line(2)).data(0).data0.poi(0)).data(0).data0.coordinate.Y - _
          m_poi(m_lin(temp_line(2)).data(0).data0.poi(1)).data(0).data0.coordinate.Y
     Call set_point_coordinate(temp_point(0).no, t_coord, False)
     Call get_new_char(temp_point(0).no)
     record_0.data0.condition_data.condition_no = 0
     Call add_point_to_line(temp_point(0).no, temp_line(2).no, 0, False, False, 0, temp_record)
 ElseIf C_display_wenti.m_point_no(num, 11) = m_lin(temp_line(2)).data(0).data0.poi(1) Then
  last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
  ' Call init_Point0(last_conditions.last_cond(1).point_no)
    temp_point(0).no = last_conditions.last_cond(1).point_no
     t_coord.X = 2 * m_poi(m_lin(temp_line(2)).data(0).data0.poi(1)).data(0).data0.coordinate.X - _
           m_poi(m_lin(temp_line(2)).data(0).data0.poi(0)).data(0).data0.coordinate.X
     t_coord.Y = 2 * m_poi(m_lin(temp_line(2)).data(0).data0.poi(1)).data(0).data0.coordinate.Y - _
          m_poi(m_lin(temp_line(2)).data(0).data0.poi(0)).data(0).data0.coordinate.Y
   Call set_point_coordinate(temp_point(0).no, t_coord, False)
   Call get_new_char(temp_point(0).no)
   record_0.data0.condition_data.condition_no = 0
   Call add_point_to_line(temp_point(0).no, temp_line(2), 0, False, False, 0, temp_record)
  End If
ElseIf C_display_wenti.m_point_no(num, 11) > 0 Then
 last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
 MDIForm1.Toolbar1.Buttons(21).Image = 33
'   Call init_Point0(last_conditions.last_cond(1).point_no)
    temp_point(0).no = last_conditions.last_cond(1).point_no
   t_coord.X = m_poi(C_display_wenti.m_point_no(num, 11)).data(0).data0.coordinate.X + _
       (m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y - _
         m_poi(C_display_wenti.m_point_no(num, 11)).data(0).data0.coordinate.Y) * 100
   t_coord.Y = m_poi(C_display_wenti.m_point_no(num, 11)).data(0).data0.coordinate.Y - _
       (m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X - _
         m_poi(C_display_wenti.m_point_no(num, 11)).data(0).data0.coordinate.X) * 100
   Call set_point_coordinate(temp_point(0).no, t_coord, False)
    last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'   Call init_Point0(last_conditions.last_cond(1).point_no)
    temp_point(1).no = last_conditions.last_cond(1).point_no
   t_coord.X = m_poi(C_display_wenti.m_point_no(num, 11)).data(0).data0.coordinate.X - _
       (m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y - _
         m_poi(C_display_wenti.m_point_no(num, 11)).data(0).data0.coordinate.Y) * 100
   t_coord.Y = m_poi(C_display_wenti.m_point_no(num, 11)).data(0).data0.coordinate.Y + _
       (m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X - _
         m_poi(C_display_wenti.m_point_no(num, 11)).data(0).data0.coordinate.X) * 100
   Call set_point_coordinate(temp_point(1).no, t_coord, False)
   Call get_new_char(temp_point(0).no)
   Call get_new_char(temp_point(1).no)
   temp_line(0) = line_number(C_display_wenti.m_point_no(num, 11), _
                              temp_point(0).no, pointapi0, pointapi0, _
                              depend_condition(point_, C_display_wenti.m_point_no(num, 11)), _
                              depend_condition(point_, temp_point(0).no), _
                              condition, condition_color, 0, 0)
      record_0.data0.condition_data.condition_no = 0
      Call add_point_to_line(temp_point(1).no, temp_line(0).no, 0, False, False, 0, temp_record)
   temp_line(0) = line_number(C_display_wenti.m_point_no(num, 11), _
                              temp_point(0).no, pointapi0, pointapi0, _
                              depend_condition(point_, C_display_wenti.m_point_no(num, 11)), _
                              depend_condition(point_, temp_point(0).no), _
                              condition, condition_color, 0, 0)
   temp_line(1) = line_number(C_display_wenti.m_point_no(num, 0), _
                              C_display_wenti.m_point_no(num, 1), _
                              pointapi0, pointapi0, _
                              depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                              depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                              condition, condition_color, 0, 0)
draw_picture1_mark_331:
 event_statue = wait_for_draw_point
    While event_statue = wait_for_draw_point
     DoEvents
    Wend
  If event_statue = draw_point_down Or _
     event_statue = draw_point_move Or _
        event_statue = draw_point_up Then 'mouse_type <> 1 Then
   t_coord = input_coord
    'temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark_331
 End If
      input_point_type% = read_inter_point(t_coord, ele1, _
                           ele2, temp_point(1).no, True)
         Call set_point_no_reduce(temp_point(1).no, 0)
If input_point_type% = new_point_on_line And ele1.no = temp_line(0) Then
 'Call draw_t_line(0)
Call set_line_visible(temp_line(0), 1)
temp_line(0) = line_number(C_display_wenti.m_point_no(num, 0), temp_point(1).no, _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                           depend_condition(point_, temp_point(1).no), _
                           condition, condition_color, 1, 0)
    Call remove_point(temp_point(0).no, no_display, 0)
   Call vertical_line(temp_line(0), temp_line(1), True, True)
    Call C_display_wenti.set_m_point_no(num, temp_point(1).no, 3, True)
    'call set_point_name(C_display_wenti.m_point_no(num,3), _
    '        C_display_wenti.m_condition(3)
    If Abs(m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.Y - _
               m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.X) > 0 Then
      A! = CSng(m_poi(C_display_wenti.m_point_no(num, 3)).data(0).data0.coordinate.X - _
                m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X) / _
                 (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y - _
                   m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y)
    Else
      A! = CSng(m_poi(C_display_wenti.m_point_no(num, 3)).data(0).data0.coordinate.Y - _
                m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y) / _
                 (m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X - _
                   m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X)
    End If
     Call C_display_wenti.set_m_point_no(num, Int(A! * 1000), 9, False)
      'Call put_name(temp_point(1))
Else
  If input_point_type% <> exist_point Then
   Call remove_point(temp_point(1).no, display, 0)
  End If
     GoTo draw_picture1_mark_331
End If
Else
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
MDIForm1.Toolbar1.Buttons(21).Image = 33
' Call init_Point0(last_conditions.last_cond(1).point_no)
  temp_point(0).no = last_conditions.last_cond(1).point_no
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
MDIForm1.Toolbar1.Buttons(21).Image = 33
' Call init_Point0(last_conditions.last_cond(1).point_no)
  temp_point(1).no = last_conditions.last_cond(1).point_no
j% = right_triangle1(m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate, _
   m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate, _
    m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.radii, _
     m_poi(temp_point(0).no).data(0).data0.coordinate.X, _
      m_poi(temp_point(0).no).data(0).data0.coordinate.Y, _
       m_poi(temp_point(1).no).data(0).data0.coordinate.X, _
        m_poi(temp_point(1).no).data(0).data0.coordinate.Y)
  'Call C_display_picture.m_BPset(Draw_form, m_poi(temp_point(0)).data(0).data0.coordinate.X, _
                                                 m_poi(temp_point(0)).data(0).data0.coordinate.Y, "", 12)
  'Call C_display_picture.m_BPset(Draw_form, m_poi(temp_point(1)).data(0).data0.coordinate.X, _
                                                 m_poi(temp_point(1)).data(0).data0.coordinate.Y, "", 12)
   Draw_form.Line (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X, _
                    m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y)- _
                     (m_poi(temp_point(0).no).data(0).data0.coordinate.X, _
                       m_poi(temp_point(0).no).data(0).data0.coordinate.Y), _
                        QBColor(fill_color)
   Draw_form.Line (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X, _
                    m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y)- _
                     (m_poi(temp_point(1).no).data(0).data0.coordinate.X, _
                      m_poi(temp_point(1).no).data(0).data0.coordinate.Y), _
                       QBColor(fill_color)
draw_picture1_mark_33:
 event_statue = wait_for_draw_point

     While event_statue = wait_for_draw_point
     DoEvents
      Wend
 If event_statue = draw_point_down Or _
     event_statue = draw_point_move Or _
        event_statue = draw_point_up Then 'mouse_type <> 1 Then
    t_coord = input_coord
    'temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark_33
End If
      input_point_type% = read_inter_point(t_coord, ele1, _
                                         ele2, temp_point(2).no, True)
        Call set_point_no_reduce(temp_point(2).no, 0)
If input_point_type% = exist_point And _
  (temp_point(0).no = temp_point(2).no Or _
     temp_point(1).no = temp_point(2).no) Then
 Draw_form.Line (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X, _
   m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y)- _
    (m_poi(temp_point(0).no).data(0).data0.coordinate.X, m_poi(temp_point(0).no).data(0).data0.coordinate.Y), _
      QBColor(fill_color)
 Draw_form.Line (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X, _
   m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y)- _
    (m_poi(temp_point(1).no).data(0).data0.coordinate.X, m_poi(temp_point(1).no).data(0).data0.coordinate.Y), _
      QBColor(fill_color)
  'Call C_display_picture.m_BPset(Draw_form, m_poi(temp_point(0)).data(0).data0.coordinate.X, _
              m_poi(temp_point(0)).data(0).data0.coordinate.Y, "", 12)
  'Call C_display_picture.m_BPset(Draw_form, m_poi(temp_point(1)).data(0).data0.coordinate.X, _
              m_poi(temp_point(1)).data(0).data0.coordinate.Y, "", 12)
  If temp_point(0).no = temp_point(2).no Then
          Call remove_point(temp_point(1).no, no_display, 0)
    Call C_display_wenti.set_m_point_no(num, 1, 10, False)
    Call C_display_wenti.set_m_point_no(num, temp_point(0).no, 3, True)
  Else
    Call remove_point(temp_point(0).no, no_display, 0)
    Call C_display_wenti.set_m_point_no(num, 2, 10, False)
    Call C_display_wenti.set_m_point_no(num, temp_point(1).no, 3, True)
  End If
    'poi(C_display_wenti.m_point_no(num,3)).data(0).data0.name = C_display_wenti.m_condition(3)
     i% = line_number(C_display_wenti.m_point_no(num, 0), _
                      C_display_wenti.m_point_no(num, 3), _
                      pointapi0, pointapi0, _
                      depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                      depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                      condition, condition_color, 1, 0)
     j% = line_number(C_display_wenti.m_point_no(num, 1), _
                      C_display_wenti.m_point_no(num, 3), _
                      pointapi0, pointapi0, _
                      depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                      depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                      condition, condition_color, 0, 0)
             Call vertical_line(i%, j%, True, True)
     Call set_point_visible(C_display_wenti.m_point_no(num, 3), 1, False)
     'Call draw_point(Draw_form, poi(C_display_wenti.m_point_no(num,3)), 0, display)
Else
 If input_point_type% <> exist_point Then
  Call remove_point(temp_point(2).no, display, 0)
 End If
 GoTo draw_picture1_mark_33
End If
End If
If C_display_wenti.m_point_no(num, 8) = 0 Then
   Call C_display_wenti.set_m_point_no(num, _
         m_circle_number(1, C_display_wenti.m_point_no(num, 1), pointapi0, _
         C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, 1, 1, _
          condition, condition_color, True), 8, False)
  Call add_point_to_m_circle(C_display_wenti.m_point_no(num, 8), _
             C_display_wenti.m_point_no(num, 11), record0, 255)
End If
End Sub

Public Sub draw_picture_31(ByVal num As Integer, ByVal no_reduce As Byte)
'-31 在□□上取一点□使得□□＝□□
Dim i%, j%, k%
Dim ele1 As condition_type
Dim ele2 As condition_type
Dim A!
For i% = 0 To 1
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
     C_display_wenti.m_condition(num, i%)) Then
  Exit Sub
End If
Next i%

For i% = 3 To 5
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
    C_display_wenti.m_condition(num, i%)) Then
     Exit Sub
End If
Next i%
    m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree = 1
 Call C_display_wenti.set_m_point_no(num, _
        C_display_picture.m_circle.Count, 8, False)
 Call m_circle_number(1, C_display_wenti.m_point_no(num, 3), pointapi0, _
         0, 0, 0, 0, C_display_wenti.m_point_no(num, 4), _
             C_display_wenti.m_point_no(num, 5), 1, 1, aid_condition, fill_color, True)
 Call C_display_wenti.set_m_point_no(num, _
       line_number(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   pointapi0, pointapi0, _
                   depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                   depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                   condition, condition_color, 1, 0), 7, False)
      Draw_form.DrawStyle = 2
  'Draw_form.Circle (m_Circ(C_display_wenti.m_point_no(num,8)).data(0).data0.c_coord.X, _
     m_Circ(C_display_wenti.m_point_no(num,8)).data(0).data0.c_coord.Y), _
      m_Circ(C_display_wenti.m_point_no(num,8)).data(0).data0.radii , QBColor(fill_color)
     Draw_form.DrawStyle = 0
draw_picture1_mark_31:
 event_statue = wait_for_draw_point

     While event_statue = wait_for_draw_point
     DoEvents
      Wend
  If event_statue = draw_point_down Or _
     event_statue = draw_point_move Or _
        event_statue = draw_point_up Then 'mouse_type <> 1 Then
    t_coord = input_coord
    'temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark_31
 End If
      input_point_type% = read_inter_point(t_coord, ele1, _
                                    ele2, temp_point(1).no, True)
        Call set_point_no_reduce(temp_point(0).no, 0)
If (input_point_type% = new_point_on_line_circle12 Or _
      input_point_type% = new_point_on_line_circle21) And _
         ele1.no = C_display_wenti.m_point_no(num, 7) And _
          ele2.no = C_display_wenti.m_point_no(num, 8) Then
  Call C_display_wenti.set_m_point_no(num, temp_point(1).no, 2, True)
 If input_point_type% = new_point_on_line_circle12 Then
  Call C_display_wenti.set_m_point_no(num, 1, 10, False)
 ElseIf input_point_type% = new_point_on_line_circle21 Then
  Call C_display_wenti.set_m_point_no(num, 2, 10, False)
 Else
  Call C_display_wenti.set_m_point_no(num, 0, 10, False)
 End If
  'poi(C_display_wenti.m_point_no(num,2)).data(0).data0.name = C_display_wenti.m_condition(num,2)
If input_point_type% = new_point_on_line_circle12 Then
  Call C_display_wenti.set_m_point_no(num, 1, 10, False)
ElseIf input_point_type% = new_point_on_line_circle21 Then
  Call C_display_wenti.set_m_point_no(num, 2, 10, False)
End If
      Draw_form.DrawStyle = 2
  Draw_form.Circle (m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.c_coord.X, _
     m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.c_coord.Y), _
      m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.radii, _
        QBColor(fill_color)
     Draw_form.DrawStyle = 0

'Call put_name(C_display_wenti.m_point_no(num,2))
Else
 If input_point_type% <> exist_point Then
  Call remove_point(temp_point(2).no, display, 0)
 End If
 GoTo draw_picture1_mark_31
End If
'End If
End Sub

Public Sub draw_picture_30(ByVal num As Integer, ByVal no_reduce As Byte)
'在⊙□[down\\(_)]上取一点□使得□□＝□□
Dim i%, j%, k%, c%
Dim ele1 As condition_type
Dim ele2 As condition_type
Dim temp_x&, temp_y&
Dim A!
Dim t_c As circle_data_type
For i% = 0 To 1
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
   C_display_wenti.m_condition(num, i%)) Then
    Exit Sub
End If
Next i%
For i% = 3 To 5
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
    C_display_wenti.m_condition(num, i%)) Then
    Exit Sub
End If
   If m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree > 2 And i% <> 0 Then
    m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree = _
     m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree - 3
   End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
c% = m_circle_number(1, C_display_wenti.m_point_no(num, 0), pointapi0, _
                     C_display_wenti.m_point_no(num, 1), 0, 0, 0, _
                      0, 0, 1, 1, aid_condition, fill_color, True)
 Call C_display_wenti.set_m_point_no(num, _
       m_circle_number(1, C_display_wenti.m_point_no(num, 0), pointapi0, _
             C_display_wenti.m_point_no(num, 1), 0, 0, 0, 0, 0, 1, 1, condition, _
                condition_color, True), 9, False)
Draw_form.DrawStyle = 2
'  Draw_form.Circle (m_Circ(C_display_wenti.m_point_no(num,8)).data(0).data0.c_coord.X, _
     m_Circ(C_display_wenti.m_point_no(num,8)).data(0).data0.c_coord.Y), _
      m_Circ(C_display_wenti.m_point_no(num,8)).data(0).data0.radii , QBColor(fill_color)
     Draw_form.DrawStyle = 0
draw_picture1_mark_30:
 event_statue = wait_for_draw_point
     While event_statue = wait_for_draw_point
     DoEvents
      Wend
  If event_statue = draw_point_down Or _
     event_statue = draw_point_move Or _
        event_statue = draw_point_up Then 'mouse_type <> 1 Then
   temp_x& = input_coord.X
    temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark_30
 End If
      input_point_type% = read_inter_point(t_coord, ele1, _
                      ele2, temp_point(1).no, 0, True)
        Call set_point_no_reduce(temp_point(0).no, 0)
If (input_point_type% = new_point_on_circle_circle12 Or _
       input_point_type% = new_point_on_circle_circle21) And _
         ((ele1.no = C_display_wenti.m_point_no(num, 9) And _
          ele2.no = C_display_wenti.m_point_no(num, 8)) Or _
           (ele1.no = C_display_wenti.m_point_no(num, 8) And _
          ele2.no = C_display_wenti.m_point_no(num, 9))) Then
 Call C_display_wenti.set_m_point_no(num, temp_point(1).no, 2, True)
  'poi(C_display_wenti.m_point_no(num,2)).data(0).data0.name = _
    C_display_wenti.m_condition (2)
If input_point_type% = new_point_on_circle_circle12 Then
 Call C_display_wenti.set_m_point_no(num, 1, 10, False)
ElseIf input_point_type% = new_point_on_circle_circle21 Then
 Call C_display_wenti.set_m_point_no(num, 2, 10, False)
End If
      Draw_form.DrawStyle = 2
  Draw_form.Circle (m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.c_coord.X, _
     m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.c_coord.Y), _
      m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.radii, QBColor(fill_color)
     Draw_form.DrawStyle = 0
  'Call put_name(C_display_wenti.m_point_no(num,2))
Else
 If input_point_type% <> exist_point Then
  Call remove_point(temp_point(2).no, display, 0)
 End If
 GoTo draw_picture1_mark_30
End If
'End If
End Sub


Public Sub draw_picture_41(ByVal num%, ByVal no_reduce As Byte)
Dim v$
Dim ang(1) As Integer
Dim i%
For i% = 0 To 5
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
       C_display_wenti.m_condition(num, i%)) Then
   Exit Sub
End If
   If m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree > 2 Then
    m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree = _
     m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree - 3
   End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
Call line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                 condition, condition_color, 1, 0)
Call line_number(C_display_wenti.m_point_no(num, 1), _
                 C_display_wenti.m_point_no(num, 2), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                 condition, condition_color, 1, 0)
Call line_number(C_display_wenti.m_point_no(num, 3), _
                 C_display_wenti.m_point_no(num, 4), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                 condition, condition_color, 1, 0)
Call line_number(C_display_wenti.m_point_no(num, 4), _
                 C_display_wenti.m_point_no(num, 5), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 4)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 5)), _
                 condition, condition_color, 1, 0)
End Sub

Public Sub draw_mid_point(ByVal p1%, ByVal p2%, p3%, is_draw As Boolean)
Dim tl%
 tl% = line_number(p1%, p2%, pointapi0, pointapi0, _
                   depend_condition(point_, p1%), _
                   depend_condition(point_, p2%), _
                   condition, condition_color, 1, 0)
 t_coord = divide_POINTAPI_by_number(add_POINTAPI(m_poi(p1%).data(0).data0.coordinate, _
                   m_poi(p2%).data(0).data0.coordinate), 2)
 Call set_point_coordinate(p3%, t_coord, False)
    m_poi(p3%).data(0).degree = 0
Call set_point_visible(p3%, 1, False)
record_0.data0.condition_data.condition_no = 0
Call add_point_to_line(p3%, tl%, 0, display, is_draw, 0, temp_record)
 ' Call draw_point(Draw_form, poi(p3%), 0, display)

End Sub
Public Sub draw_picture_32_3(ByVal num As Integer, ByVal no_reduce As Byte)
'与⊙□[down\\(_)]相切于点□的切线交直线□□于□
Dim i%, j%, k%, t_point1%
Dim ele1 As condition_type
Dim ele2 As condition_type
'Dim temp_x&, temp_y&
For i% = 0 To 4
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
      C_display_wenti.m_condition(num, i%)) Then
  Exit Sub
End If
   If m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree > 2 Then
    m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree = _
     m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree - 3
   End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
i% = m_circle_number(1, C_display_wenti.m_point_no(num, 0), pointapi0, _
                        C_display_wenti.m_point_no(num, 1), 0, 0, 0, 0, _
                        0, 1, 1, condition, condition_color, True)
'If open_record Then
'   last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'   c_display_wenti.m_point_no(num,5) = last_conditions.last_cond(1).point_no
'    poi(last_conditions.last_cond(1).point_no).data(0).data0.name = c_display_wenti.m_condition(num,5)
'     poi(last_conditions.last_cond(1).point_no).data(0).data0.coordinate.X = temp_record_poi(last_conditions.last_cond(1).point_no, 0)
'      poi(last_conditions.last_cond(1).point_no).data(0).data0.coordinate.Y = temp_record_poi(last_conditions.last_cond(1).point_no, 1)
 '      poi(last_conditions.last_cond(1).point_no).data(0).data0.visible = 1
 '       Call draw_point(Draw_form, poi(last_conditions.last_cond(1).point_no).data(0), display)
' j% = line_number(c_display_wenti.m_point_no(num,2), c_display_wenti.m_point_no(num,0), _
        condition, no_display)
' k% = line_number(c_display_wenti.m_point_no(num,2), c_display_wenti.m_point_no(num,5), _
       condition, display)
'          Call vertical_line(k%, j%, True)
'If c_display_wenti.m_no = -32 Then
' j% = line_number(c_display_wenti.m_point_no(num,3), c_display_wenti.m_point_no(num,4), _
   condition, display)
'Call add_point_to_line(c_display_wenti.m_point_no(num,5), j%, True, _
    display)
'Else
'j% =C_display_picture.m_circle_number(c_display_wenti.m_point_no(num,3), c_display_wenti.m_point_no(num,4))
'record_0 = record0
'Call add_point_to_circle(c_display_wenti.m_point_no(num,5), j%, record_0, 0)
'End If
'Else
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
MDIForm1.Toolbar1.Buttons(21).Image = 33
  ' Call init_Point0(last_conditions.last_cond(1).point_no)
temp_point(0).no = last_conditions.last_cond(1).point_no
t_coord.X = m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.X + _
 (m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y - _
    m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y) * 100
t_coord.Y = m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.Y - _
 (m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X - _
    m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X) * 100
Call set_point_coordinate(temp_point(0).no, t_coord, False)
temp_line(0) = line_number(C_display_wenti.m_point_no(num, 2), temp_point(0).no, _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                           depend_condition(point_, temp_point(0).no), _
                           condition, condition_color, 0, 0)
temp_line(1) = line_number(C_display_wenti.m_point_no(num, 0), _
                           C_display_wenti.m_point_no(num, 1), _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                           condition, condition_color, 0, 0)
                   Call vertical_line(temp_line(0), temp_line(1), True, True)
If C_display_wenti.m_no(num) = -32 Then
  temp_line(2) = line_number(C_display_wenti.m_point_no(num, 3), _
                             C_display_wenti.m_point_no(num, 4), _
                             pointapi0, pointapi0, _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 4)), _
                             condition, condition_color, 1, 0)
Else
 Call C_display_wenti.set_m_point_no(num, _
       m_circle_number(1, C_display_wenti.m_point_no(num, 3), pointapi0, _
         C_display_wenti.m_point_no(num, 4), 0, 0, 0, 0, 0, _
          1, 1, condition, condition_color, True), 9, False)
End If
draw_picture1_mark_32:
 event_statue = wait_for_draw_point
   While event_statue = wait_for_draw_point
     DoEvents
      Wend
 If event_statue = draw_point_down Or _
     event_statue = draw_point_move Or _
        event_statue = draw_point_up Then 'mouse_type <> 1 Then
   t_coord = input_coord
    'temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark_32
End If
      input_point_type% = read_inter_point(t_coord, ele1, _
                                     ele2, temp_point(1).no, True)
         Call set_point_no_reduce(temp_point(1).no, 0)
If C_display_wenti.m_no(num) = -32 And input_point_type% = interset_point_line_line And (ele1.no = temp_line(0) Or _
    ele1.no = temp_line(2)) And (ele2.no = temp_line(2) Or _
     ele2.no = temp_line(0)) Then
Call draw_tangent_line(0)
Call draw_tangent_line(1)
 Call remove_point(temp_point(0).no, no_display, 0)
 Call set_line_visible(temp_line(0), 1)
temp_line(2) = line_number(C_display_wenti.m_point_no(num, 2), temp_point(1).no, _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                           depend_condition(point_, temp_point(1).no), _
                           condition, condition_color, 1, 0)
    Call C_display_wenti.set_m_point_no(num, temp_point(1).no, 5, False)
     'poi(C_display_wenti.m_point_no(num,5)).data(0).data0.name = _
       C_display_wenti.m_condition (5)
     'Call put_name(temp_point(1))
ElseIf C_display_wenti.m_no(num) = -3 And _
  (input_point_type% = new_point_on_line_circle12 Or _
   input_point_type% = new_point_on_line_circle21) And _
   ele1.no = temp_line(0) And ele2.no = C_display_wenti.m_point_no(num, 9) Then
Call draw_tangent_line(0)
Call draw_tangent_line(1)
 Call remove_point(temp_point(0).no, no_display, 0)
 Call set_line_visible(temp_line(0), 1)
temp_line(0) = line_number(C_display_wenti.m_point_no(num, 2), temp_point(1).no, _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                           depend_condition(point_, temp_point(1).no), _
                           condition, condition_color, 1, 0)
    Call C_display_wenti.set_m_point_no(num, temp_point(1).no, 5, False)
     'poi(C_display_wenti.m_point_no(num,5)).data(0).data0.name = _
         C_display_wenti.m_condition (5)
    ' Call put_name(temp_point(1))
   If input_point_type% = new_point_on_line_circle12 Then
    Call C_display_wenti.set_m_point_no(num, 1, 7, False)
   Else
    Call C_display_wenti.set_m_point_no(num, 2, 7, False)
   End If
Else
 If input_point_type% <> exist_point Then
  Call remove_point(temp_point(1).no, display, 0)
  End If
     GoTo draw_picture1_mark_32
 End If
'End If
    last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
    MDIForm1.Toolbar1.Buttons(21).Image = 33
'     Call init_Point0(last_conditions.last_cond(1).point_no)
     Call C_display_wenti.set_m_point_no(num, _
           last_conditions.last_cond(1).point_no, 15, False)
      Call get_new_char(last_conditions.last_cond(1).point_no)
 t_coord.X = 2 * m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.X - _
       m_poi(C_display_wenti.m_point_no(num, 5)).data(0).data0.coordinate.X
 t_coord.Y = 2 * m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.Y - _
       m_poi(C_display_wenti.m_point_no(num, 5)).data(0).data0.coordinate.Y
        Call set_point_coordinate(C_display_wenti.m_point_no(num, 15), _
                              t_coord, False)
record_0.data0.condition_data.condition_no = 0
Call add_point_to_line(last_conditions.last_cond(1).point_no, _
   line_number0(C_display_wenti.m_point_no(num, 2), _
   C_display_wenti.m_point_no(num, 5), 0, 0), 0, _
    no_display, True, 0, temp_record)
End Sub
Public Sub remove_brac(ByVal w_n%, ByVal s%, ByVal n%)
Dim p%, i%
If C_display_wenti.m_condition(w_n%, s%) = "(" And _
    C_display_wenti.m_condition(w_n%, n%) = ")" Then
p% = 1
 For i% = s% + 1 To n% - 1
  If C_display_wenti.m_condition(w_n%, n%) = ")" Then
   p% = p% - 1
  ElseIf C_display_wenti.m_condition(w_n%, n%) = "(" Then
   p% = p% + 1
  End If
   If p% = 0 Then
    Exit Sub
   End If
 Next i%
      s% = s% + 1
       n% = n% - 1
End If
End Sub
Public Sub read_multi_item(ByVal w_n%, ByVal s%, ByVal n%, _
            item As item0_data_type, para As String, ty As Byte, set_v As Byte)  'para As String, p1%, p2%, p3%, p4%, sig As String)
'从输入语句的S% 到N% 读出一项
Dim i%, j%, k%, l%, m%, t%
Dim ts$
Dim tp$
Dim tp1$
Dim A(3) As Integer
Dim sig_(3) As String
Dim p1_(3) As Integer
Dim p2_(3) As Integer
Dim it(3) As Integer
Dim total%
Dim lbrac_n(3) As Integer
Dim rbrac_n(3) As Integer
Dim divide_sig As Integer
Dim time_sig As Integer
Dim p(11) As Integer
Dim tl(5) As Integer
Dim si As String
Dim po(3) As Integer
Dim sig As String
Dim para_ As String
Dim v As String
Dim sig_ty As Byte
Dim t_item As item0_data_type
Dim t_para As String
Dim t_ord(1) As Integer
Dim ord As Integer
'para = "1"
'Call remove_brac(w_n%, s%, n%)
'去括号
For i% = s% To n%
If C_display_wenti.m_condition(w_n%, i%) >= "A" And _
      C_display_wenti.m_condition(w_n%, i%) <= "Z" Then
 GoTo read_multi_item_start
ElseIf C_display_wenti.m_condition(w_n%, i%) = Chr(12) Then '"," Then
ts$ = C_display_wenti.m_condition(w_n%, i%)
 A(0) = Abs(angle_number(C_display_wenti.m_condition(w_n%, i% + 1), _
      C_display_wenti.m_condition(w_n%, i% + 2), _
        C_display_wenti.m_condition(w_n%, i% + 3), "", 0))
   item.sig = "~"
   item.poi(0) = A(0)
   item.poi(1) = -6
   item.para(0) = "1"
   item.para(1) = "1"
   If i% = s% Then
    para = "1"
   Else
    If C_display_wenti.m_condition(w_n%, i% - 1) = "*" Then
     i% = i% - 1
    End If
    'Call read_number_from_wenti(w_n%, s%, i%, para)
   End If
   Exit Sub
Else
ts$ = ts$ + C_display_wenti.m_condition(w_n%, i%)
End If
Next i%
read_multi_item_mark2:
If InStr(1, ts$, "sin", 0) > 0 Then
 GoTo read_multi_item_start
ElseIf InStr(1, ts$, "cos", 0) > 0 Then
 GoTo read_multi_item_start
ElseIf InStr(1, ts$, "tan", 0) > 0 Then
 GoTo read_multi_item_start
ElseIf InStr(1, ts$, "ctan", 0) > 0 Then
 GoTo read_multi_item_start
ElseIf ts$ = Chr(12) Then ' "," Then
 GoTo read_multi_item_start
Else
'Call read_number_from_wenti(w_n%, s, n%, para)
item.sig = "F"
Exit Sub
End If
read_multi_item_start:
ord = 1
If ty = 1 Then
para = "1"
End If
If (C_display_wenti.m_condition(w_n%, s%) > "9" And _
      C_display_wenti.m_condition(w_n%, s%) < "a") Or _
         C_display_wenti.m_condition(w_n%, s%) < "0" Or _
          C_display_wenti.m_condition(w_n%, s%) > "z" Then '无系数
If ty = 1 Then '第一次
para = "1"
End If
 If C_display_wenti.m_condition(w_n%, s%) = "(" And _
      C_display_wenti.m_condition(w_n%, n%) = ")" Then '去括号
   Call read_multi_item(w_n%, s% + 1, n% - 1, item, para, 0, set_v)
      Exit Sub
 Else
 t% = 0
 For i% = s% To n% '(AB/CD)^2
  If C_display_wenti.m_condition(w_n%, i%) = "(" Then
   lbrac_n(0) = i%
   If t% = 0 Then
   l% = i%
   End If
    t% = t% + 1 '括号
  ElseIf C_display_wenti.m_condition(w_n%, i%) = "/" Then
    k% = i%
  ElseIf C_display_wenti.m_condition(w_n%, i%) = ")" Then
   If t% = 0 Then
    m% = i%
     If l% < k% And k% < m% Then
      If m% <= n% - 2 Then
        If C_display_wenti.m_condition(w_n%, m% + 1) = "^" Then
         ord = val(C_display_wenti.m_condition(w_n%, m% + 2))
          Call read_multi_item(w_n%, s% + 1, m% - 1, item, t_para, 0, set_v)
           para = time_string(para, t_para, True, False)
             Exit Sub
         End If
     End If
   End If
   End If
  t% = t% - 1
   If t% = 0 Then
    rbrac_n(0) = i%
   End If
  End If
 Next i%
  For i% = n% To s% Step -1 '()/3
   If C_display_wenti.m_condition(w_n%, i%) >= "0" And _
        C_display_wenti.m_condition(w_n%, i%) <= "9" Then
    tp$ = C_display_wenti.m_condition(w_n%, i%) + tp$
   ElseIf C_display_wenti.m_condition(w_n%, i%) = "/" Then
    If tp$ <> "" Then
     l% = 0
     For j% = i% + 2 To n%
      tp1$ = C_display_wenti.m_condition(w_n%, j%)
       If (tp1$ >= "0" And tp1$ <= "9") Or tp$ = "'" Or tp1$ = LoadResString_(1460, "") Then
         tp$ = tp$ + tp1$
       ElseIf tp1$ >= "a" And tp1$ <= "z" Then
 '       call
       ElseIf tp1$ >= "A" And tp1$ <= "B" Then
        l% = l% + 1
         p1_(l%) = C_display_wenti.m_condition(w_n%, j%)
       ElseIf tp1$ = empty_char Then
        'if c_display_wenti.point_no (j%)<0
       End If
     Next j%
      para = divide_string(para, tp$, True, False)
'         Call read_multi_item(w_n%, s%, i%, item, t_para, 0)
 '         para = time_string(para, t_para, True, False)
           Call read_multi_item(w_n%, s%, i% - 1, item, para, 0, set_v)
            If item.sig = "~" And l% = 2 Then
               item.sig = "/"
               item.line_no(1) = line_number0(p1_(0), p1_(1), item.n(2), item.n(3))
               If item.n(2) < item.n(3) Then
               item.poi(2) = p1_(0)
               item.poi(3) = p1_(1)
               Else
               Call exchange_two_integer(item.n(2), item.n(3))
               item.poi(2) = p1_(1)
               item.poi(3) = p1_(0)
               End If
            End If
            Exit Sub
             
     End If
   ElseIf (C_display_wenti.m_condition(w_n%, i%) >= "A" And _
            C_display_wenti.m_condition(w_n%, i%) <= "Z") Or _
            C_display_wenti.m_condition(w_n%, i%) = ")" Then
      GoTo read_multi_item_mark10
   End If
  Next i%
read_multi_item_mark10:
  For i% = s% To n%
   If C_display_wenti.m_condition(w_n%, i%) = "/" Then
     If i% = s% Then
     item.poi(0) = 0
     item.poi(1) = 0
      Call read_element_from_wenti(w_n%, i% + 1, n, item.poi(2), item.poi(3), t_ord(0))
       item.sig = "/"
          Exit Sub
     Else
      Call read_element_from_wenti(w_n%, s%, i% - 1, item.poi(0), item.poi(1), t_ord(0))
      Call read_element_from_wenti(w_n%, i% + 1, n, item.poi(2), item.poi(3), t_ord(1))
      If t_ord(0) = t_ord(1) Then
         ord% = ord% * t_ord(0)
      End If
       If ord% >= 2 Or set_v = 1 Then
       Call set_drelation_for_add(item.poi(0), item.poi(1), item.poi(2), item.poi(3), v)
       item.poi(0) = 0
       item.poi(0) = 0
       item.poi(0) = 0
       item.poi(0) = 0
       item.sig = "F"
       For j% = 1 To ord% - 1
       para = time_string(para, v, False, False)
       Next j%
       para = time_string(para, v, True, False)
      Else
       item.sig = "/"
      End If
      Exit Sub
    End If
   ElseIf C_display_wenti.m_condition(w_n%, i%) = "*" Or _
            C_display_wenti.m_condition(w_n%, i%) = "-" Or _
             C_display_wenti.m_condition(w_n%, i%) = "+" Then
      Call read_element_from_wenti(w_n%, s%, i% - 1, item.poi(0), item.poi(1), t_ord(0))
      Call read_element_from_wenti(w_n%, i% + 1, n, item.poi(2), item.poi(3), t_ord(1))
       item.sig = C_display_wenti.m_condition(w_n%, i%)
          Exit Sub
   End If
   Next i%
   Call read_element_from_wenti(w_n%, s%, n%, item.poi(0), item.poi(1), t_ord(0))
    If t_ord(0) = 1 Then
    item.sig = "~"
    ElseIf t_ord(0) = 2 Then
     item.poi(2) = item.poi(0)
     item.poi(3) = item.poi(1)
     item.sig = "*"
    End If
      Exit Sub
     End If
Else
para = ""
For i% = s% To n%
 If C_display_wenti.m_condition(w_n%, i%) >= "0" And C_display_wenti.m_condition(w_n%, i%) <= "9" Then
  para = para + C_display_wenti.m_condition(w_n%, i%)
 Else
 If ty = 1 Then
  If para = "" Then
   para = "1"
  End If
 End If
  If C_display_wenti.m_condition(w_n%, i%) = "*" Then
    i% = i% + 1
  End If
  Call read_multi_item(w_n%, i%, n%, item, para, 0, set_v)
 '  para = time_string(para, tp$, True, False)
      Exit Sub
  End If
Next i%
End If
  item = set_item0_(item.poi(0), item.poi(1), item.poi(2), item.poi(3), item.sig, 0, 0, 0, 0, 0, 0, _
            "", "", "", 0, "", condition_data0)
'End If
 End Sub
Public Sub draw_picture38_49(ByVal num%, ByVal st%, ByVal en%, no_reduce)
'38 _~
Dim i%, j%, k%, l%, tn%, last%, last1%, no%
'Dim p(15) As Integer
Dim m(3) As String
Dim sig(3) As String
Dim para(10) As String
Dim it_para(3) As String
Dim tv As String
Dim set_tv As String
Dim it(3) As Integer
Dim para_(10) As String
Dim tv_ As String
'Dim it_(3) As Integer
Dim t_it(1) As Integer
Dim item(10) As item0_data_type
Dim item_(10) As item0_data_type
Dim temp_record As total_record_type
Dim is_zero As Boolean
Dim s_brace%, e_brace%, b%
Dim equal_no(5) As Integer
Dim last_equal_no%
Dim con_ty As Byte
Dim con_string As String
For i% = st% To en% '确定等号个数
 If C_display_wenti.m_condition(num, i%) = empty_char Then
     If i% < en% Then
      If C_display_wenti.m_condition(num, i% + 1) <> empty_char Then
       st% = i% + 1
      End If
     End If
     If i% > st% Then
      If C_display_wenti.m_condition(num, i% - 1) <> empty_char Then
       en% = i% - 1
      End If
     End If
 ElseIf C_display_wenti.m_condition(num, i%) = "=" Then
  last_equal_no% = last_equal_no% + 1
   equal_no(last_equal_no%) = i%
  End If
 Next i%
If last_equal_no% > 1 Then '多次输入
 equal_no(0) = st% - 1
 equal_no(last_equal_no% + 1) = en% + 1
 For i% = 1 To last_equal_no%
 Call draw_picture38_49(num%, equal_no(i% - 1) + 1, equal_no(i% + 1) - 1, 0)
 Next i%
 Exit Sub
Else
For i% = 0 To 3 '初始化系数
para(i%) = "0"
Next i%
If equal_no(1) > st% And equal_no(1) < en% Then   '有等号
 Call read_multi_item_(num%, st%, equal_no(1) - 1, "1", para(), item(), last1%) '等号前
 Call read_multi_item_(num%, equal_no(1) + 1, en%, "-1", para_(), item_(), last%) '等号后
    set_tv = "0"
Else
  Call read_multi_item_(num%, st%, en%, "1", para(), item(), last1%)
      set_tv = ""
End If
 If last% > 0 Then '将等号右移等号左
  is_zero = True
      For i% = 0 To last% - 1
      para(last1% + i%) = para_(i%)
      item(last1% + i%) = item_(i%)
      Next i%
 Else
      last% = last1%
 End If
End If
   temp_record.record_data.data0.condition_data.condition_no = 0 'record0
     temp_record.record_.display_no = -num
If last1% > 4 Then
   set_tv = set_new_value(item(), para(), 0, last1% - 2)
  item(0) = item(last1% - 1)
  para(0) = para(last1% - 1)
  para(1) = "0"
  para(2) = "0"
  para(3) = "0"
  last1% = 1
End If
If (item(0).poi(1) > -1 And item(0).poi(1) < -6) Or _
    (item(1).poi(1) > -1 And item(1).poi(1) < -6) Or _
     (item(2).poi(1) > -1 And item(2).poi(1) < -6) Or _
      (item(3).poi(1) > -1 And item(3).poi(1) < -6) Then
       th_chose(118).chose = 1
End If
If C_display_wenti.m_no(num) = -49 Then
 If para(0) = "0" Then
  Exit Sub
 ElseIf para(1) = "0" Then
' ??????
 ElseIf para(2) = "0" Then
  If item(0).sig = "*" And item(1).sig = "*" And _
       para(0) = time_string("-1", para(1), True, False) Then
    tn% = 0
    Call set_dpoint_pair(item(0).poi(0), item(0).poi(1), _
     item(1).poi(0), item(1).poi(1), item(1).poi(2), _
      item(1).poi(3), item(0).poi(2), item(0).poi(3), _
       item(0).n(0), item(0).n(1), item(1).n(0), item(1).n(1), _
        item(1).n(2), item(1).n(3), item(0).n(2), item(0).n(3), _
         item(0).line_no(0), item(1).line_no(0), item(1).line_no(1), _
          item(0).line_no(1), 1, temp_record, False, tn%, 0, 0, 0, True)
           'Ddpoint_pair(tn%).data(0).record.display_no = -(num + 1)
         GoTo draw_picture38_49_mark1
  ElseIf item(0).sig = "/" And item(1).sig = "/" Then
    If para(0) = time_string("-1", para(1), True, False) Then
      If item(0).poi(0) > 0 And item(0).poi(1) > 0 And _
        item(1).poi(0) > 0 And item(1).poi(1) > 0 Then
       Call set_dpoint_pair(item(0).poi(0), item(0).poi(1), _
        item(0).poi(2), item(0).poi(3), item(1).poi(0), _
         item(1).poi(1), item(1).poi(2), item(1).poi(3), _
          item(0).n(0), item(0).n(1), item(0).n(2), item(0).n(3), _
           item(1).n(0), item(1).n(1), item(1).n(2), item(1).n(3), _
            item(0).line_no(0), item(0).line_no(1), item(1).line_no(0), item(1).line_no(1), _
             0, temp_record, False, 0, 0, 0, no_reduce, True)
           GoTo draw_picture38_49_mark1
      ElseIf item(0).poi(0) = 0 And item(0).poi(1) = 0 And _
             item(1).poi(0) = 0 And item(1).poi(1) = 0 Then
       Call set_equal_dline(item(0).poi(2), item(0).poi(3), _
         item(1).poi(2), item(1).poi(3), item(0).n(2), item(0).n(3), _
          item(1).n(2), item(1).n(3), item(0).line_no(1), item(1).line_no(1), _
           0, temp_record, 0, 0, 0, 0, no_reduce, True)
         GoTo draw_picture38_49_mark1
      End If
    Else 'para(0)<>-para(1)
      If item(0).poi(0) = 0 And item(0).poi(1) = 0 And _
             item(1).poi(0) = 0 And item(1).poi(1) = 0 Then
        Call set_Drelation(item(1).poi(2), item(1).poi(3), _
             item(0).poi(2), item(0).poi(3), item(1).n(2), _
              item(1).n(3), item(0).n(2), item(0).n(3), _
               item(1).line_no(1), item(0).line_no(1), _
                divide_string(time_string("-1", para(1), False, False), para(0), True, False), _
                 temp_record, 0, 0, 0, 0, no_reduce, True)
                  GoTo draw_picture38_49_mark1
      End If
    End If
  ElseIf item(0).sig = "~" And item(1).sig = "~" Then
    If item(0).poi(1) = -6 And item(1).poi(1) = -7 Then
     Call set_three_angle_value(item(0).poi(0), item(1).poi(0), 0, para(0), para(1), "0", "0", _
        0, temp_record, 0, 0, 0, 0, 0, 0, False)
          GoTo draw_picture38_49_mark1
    End If
    If para(0) = time_string("-1", para(1), True, False) Then
      Call set_equal_dline(item(0).poi(0), item(0).poi(1), _
       item(1).poi(0), item(1).poi(1), item(0).n(0), item(0).n(1), _
        item(1).n(0), item(1).n(1), item(0).line_no(0), _
         item(1).line_no(0), 0, temp_record, 0, 0, 0, 0, no_reduce, True)
         GoTo draw_picture38_49_mark1
    Else
     Call set_Drelation(item(0).poi(0), item(0).poi(1), _
        item(1).poi(0), item(1).poi(1), item(0).n(0), _
         item(0).n(1), item(1).n(0), item(1).n(1), _
          item(0).line_no(0), item(1).line_no(0), divide_string( _
           time_string("-1", para(1), False, False), para(0), True, False), _
         temp_record, 0, 0, 0, 0, no_reduce, True)
          GoTo draw_picture38_49_mark1
    End If
  ElseIf item(0).sig = "~" And item(1).sig = "F" Then
   Call input_data_from_item(item(0), _
     divide_string(time_string("-1", para(1), False, False), para(0), True, False), _
       temp_record)
   GoTo draw_picture38_49_mark1
 '  ElseIf item(0).poi(0) = -1 Then 'End If
  ' elseif
  '  Call set_line_value(item(0).poi(0), item(0).poi(1), _
  '   divide_string(time_string("-1", para(1), False, False), para(0), True, False), _
  '    item(0).n(0), item(0).n(1), item(0).line_no(0), _
  '     temp_record, 0, no_reduce)
  '       GoTo draw_picture38_49_mark1
  ElseIf item(0).sig = "F" And item(1).sig = "~" Then
     Call input_data_from_item(item(1), _
     divide_string(time_string("-1", para(0), False, False), para(1), True, False), _
             temp_record)
         GoTo draw_picture38_49_mark1
  ElseIf item(0).sig = "/" And item(1).sig = "F" Then
     If item(0).poi(0) > 0 And item(0).poi(1) > 0 Then
      Call set_Drelation(item(0).poi(0), item(0).poi(1), _
       item(0).poi(2), item(0).poi(3), item(0).n(0), _
        item(0).n(1), item(0).n(2), item(0).n(3), _
         item(0).line_no(0), item(0).line_no(1), divide_string( _
        time_string("-1", para(1), False, False), para(0), True, False), temp_record, 0, _
         0, 0, 0, no_reduce, True)
          GoTo draw_picture38_49_mark1
     Else
      Call set_line_value(item(0).poi(2), item(0).poi(3), _
        divide_string(time_string("-1", para(0), False, False), para(1), True, False), _
          item(0).n(2), item(0).n(3), item(0).line_no(1), _
           temp_record.record_data, 0, no_reduce, False)
          GoTo draw_picture38_49_mark1
    End If
  ElseIf item(0).sig = "F" And item(1).sig = "/" Then
     If item(1).poi(0) > 0 And item(1).poi(1) > 0 Then
      Call set_Drelation(item(1).poi(0), item(1).poi(1), _
       item(1).poi(2), item(1).poi(3), item(1).n(0), _
        item(1).n(1), item(1).n(2), item(1).n(3), _
         item(1).line_no(0), item(1).line_no(1), divide_string( _
        time_string("-1", para(0), False, False), para(1), True, False), temp_record, 0, _
         0, 0, 0, no_reduce, True)
          GoTo draw_picture38_49_mark1
     Else
      Call set_line_value(item(1).poi(2), item(1).poi(3), _
        divide_string(time_string("-1", para(1), False, False), para(0), True, False), _
         item(1).n(2), item(1).n(3), item(1).line_no(1), _
          temp_record.record_data, 0, no_reduce, False)
         GoTo draw_picture38_49_mark1
     End If
  End If
 ElseIf para(3) = "0" Then
   If item(0).sig = "~" And item(1).sig = "~" And item(2).sig = "~" Then
    Call set_three_line_value(item(0).poi(0), item(0).poi(1), _
     item(1).poi(0), item(1).poi(1), item(2).poi(0), item(2).poi(1), _
      item(0).n(0), item(0).n(1), item(1).n(0), item(1).n(1), item(2).n(0), _
       item(2).n(1), item(0).line_no(0), item(1).line_no(0), item(2).line_no(0), _
        para(0), para(1), para(2), "0", temp_record, 0, no_reduce, 0)
         GoTo draw_picture38_49_mark1
   ElseIf item(0).sig = "~" And item(1).sig = "~" And (item(2).sig = "F" Or _
     item(2).sig = empty_char) Then
    Call set_two_line_value(item(0).poi(0), item(0).poi(1), _
     item(1).poi(0), item(1).poi(1), item(0).n(0), item(0).n(1), _
      item(1).n(0), item(1).n(1), item(0).line_no(0), item(1).line_no(1), _
       para(0), para(1), time_string("-1", para(2), True, False), temp_record, 0, _
        no_reduce)
         GoTo draw_picture38_49_mark1
   ElseIf item(0).sig = "~" And (item(1).sig = "F" Or item(1).sig = empty_char) And _
      item(2).sig = "~" Then
    Call set_two_line_value(item(0).poi(0), item(0).poi(1), _
     item(2).poi(0), item(2).poi(1), item(0).n(0), item(0).n(1), _
      item(2).n(0), item(2).n(1), item(0).line_no(0), item(2).line_no(0), _
       para(0), para(2), time_string("-1", para(1), True, False), _
        temp_record, 0, no_reduce)
         GoTo draw_picture38_49_mark1
   ElseIf (item(0).sig = "F" Or item(0).sig = empty_char) And item(1).sig = "~" And _
     item(2).sig = "~" Then
    Call set_two_line_value(item(1).poi(0), item(1).poi(1), _
     item(2).poi(0), item(2).poi(1), item(1).n(0), item(1).n(1), _
       item(2).n(0), item(2).n(1), item(1).line_no(0), item(2).line_no(0), _
         para(1), para(2), time_string("-1", para(0), True, False), temp_record, 0, _
          no_reduce)
         GoTo draw_picture38_49_mark1
   ElseIf item(0).sig = "/" And item(1).sig = "/" And item(2).sig = "/" Then
    If item(0).poi(0) = 0 And item(0).poi(1) = 0 And item(1).poi(0) = 0 And _
        item(1).poi(1) = 0 And item(2).poi(0) = 0 And item(2).poi(1) = 0 Then
        item(0) = set_item0_(item(2).poi(2), item(2).poi(3), _
           item(0).poi(2), item(0).poi(3), "/", _
            item(2).n(2), item(2).n(3), item(0).n(2), _
             item(0).n(3), item(2).line_no(1), item(0).line_no(1), "", "", "", 0, _
                "", temp_record.record_data.data0.condition_data)
      item(1) = set_item0_(item(2).poi(2), item(2).poi(3), item(1).poi(2), _
            item(1).poi(3), "/", 0, 0, 0, 0, 0, 0, "", "", "", 0, "", condition_data0)
      item(2) = set_item0_(0, 0, 0, 0, "", 0, 0, 0, 0, 0, 0, "", "", "", 0, "", _
                                           temp_record.record_data.data0.condition_data)
    End If
   End If
 Else
  If item(0).sig = "~" And item(1).sig = "~" And item(2).sig = "~" _
         And item(3).sig = "F" Then
   Call set_three_line_value(item(0).poi(0), item(0).poi(1), _
    item(1).poi(0), item(1).poi(1), item(2).poi(0), item(2).poi(1), _
     item(0).n(0), item(0).n(1), item(1).n(0), item(1).n(1), item(2).n(0), _
      item(2).n(1), item(0).line_no(0), item(1).line_no(0), item(2).line_no(0), _
     para(0), para(1), para(2), divide_string( _
      time_string("-1", para(3), False, False), para(0), True, False), temp_record, _
        0, no_reduce, 0)
          GoTo draw_picture38_49_mark1
  ElseIf item(0).sig = "~" And item(1).sig = "~" And item(2).sig = "F" _
         And item(3).sig = "~" Then
     Call set_three_line_value(item(0).poi(0), item(0).poi(1), _
        item(1).poi(0), item(1).poi(1), item(3).poi(0), item(3).poi(1), _
         item(0).n(0), item(0).n(1), item(1).n(0), item(1).n(1), item(3).n(0), _
          item(3).n(1), item(0).line_no(0), item(1).line_no(0), item(3).line_no(0), _
          para(0), para(1), para(3), time_string("-1", para(3), True, False), temp_record, _
            0, no_reduce, 0)
         GoTo draw_picture38_49_mark1
  ElseIf item(0).sig = "~" And item(1).sig = "F" And item(2).sig = "~" _
          And item(3).sig = "~" Then
     Call set_three_line_value(item(0).poi(0), item(0).poi(1), _
      item(2).poi(0), item(2).poi(1), item(3).poi(0), item(3).poi(1), _
       item(0).n(0), item(0).n(1), item(2).n(0), item(2).n(1), item(3).n(0), _
        item(3).n(1), item(0).line_no(0), item(2).line_no(0), item(3).line_no(0), _
        para(0), para(2), para(3), time_string("-1", para(1), True, False), temp_record, _
         0, no_reduce, 0)
          GoTo draw_picture38_49_mark1
  ElseIf item(0).sig = "F" And item(1).sig = "~" And item(2).sig = "~" _
          And item(3).sig = "~" Then
     Call set_three_line_value(item(1).poi(0), item(1).poi(1), _
       item(2).poi(0), item(2).poi(1), item(3).poi(0), item(3).poi(1), _
        item(1).n(0), item(1).n(1), item(2).n(0), item(2).n(1), item(3).n(0), _
         item(3).n(1), item(1).line_no(0), item(2).line_no(0), item(3).line_no(0), _
        para(1), para(2), para(3), time_string("-1", para(0), True, False), temp_record, _
         0, no_reduce, 0)
          GoTo draw_picture38_49_mark1
  ElseIf item(0).sig = "/" And item(1).sig = "/" And item(2).sig = "/" _
         And item(3).sig = "/" Then
     If item(0).poi(0) = 0 And item(0).poi(1) = 0 And item(1).poi(0) = 0 And _
         item(1).poi(1) = 0 And item(2).poi(0) = 0 And item(2).poi(1) And _
          item(3).poi(0) = 0 And item(3).poi(1) = 0 Then
       item(0) = set_item0_(item(3).poi(2), item(3).poi(3), _
                        item(0).poi(2), item(0).poi(3), "/", _
                        item(3).n(2), item(3).n(3), item(0).n(2), _
                         item(0).n(3), item(3).line_no(1), item(0).line_no(1), "", "", "", 0, _
                            "", temp_record.record_data.data0.condition_data)
       item(1) = set_item0_(item(3).poi(2), item(3).poi(3), _
                        item(1).poi(2), item(1).poi(3), "/", _
                         item(2).n(2), item(3).n(3), item(1).n(2), _
                          item(1).n(3), item(3).line_no(1), item(1).line_no(1), "", "", "", 0, _
                             "", temp_record.record_data.data0.condition_data)
       item(2) = set_item0_(item(3).poi(2), item(3).poi(3), _
                        item(2).poi(2), item(2).poi(3), "/", _
                         item(3).n(2), item(3).n(3), item(2).n(2), _
                          item(2).n(3), item(3).line_no(1), item(2).line_no(1), "", "", "", 0, _
                             "", temp_record.record_data.data0.condition_data)
       item(3) = set_item0_(0, 0, 0, 0, "", 0, 0, 0, 0, 0, 0, "", "", "", 0, _
                            "", temp_record.record_data.data0.condition_data)
     End If
  End If
 End If
   it(0) = item_number(item(0), it_para(0), temp_record.record_data.data0.condition_data)
   it(1) = item_number(item(1), it_para(1), temp_record.record_data.data0.condition_data)
   it(2) = item_number(item(2), it_para(2), temp_record.record_data.data0.condition_data)
   it(3) = item_number(item(3), it_para(3), temp_record.record_data.data0.condition_data)
   para(0) = time_string(para(0), it_para(0), True, False)
   para(1) = time_string(para(1), it_para(1), True, False)
   para(2) = time_string(para(2), it_para(2), True, False)
   para(3) = time_string(para(3), it_para(3), True, False)
    Call set_general_string(it(0), it(1), it(2), it(3), para(0), para(1), _
            para(2), para(3), time_string("-1", set_tv, True, False), 0, 0, 1, _
               temp_record, 0, no_reduce)
         GoTo draw_picture38_49_mark1
Else 'If C_display_wenti.m_no = 38 Then '38
 If set_tv <> "" Then
 For i% = 0 To 3
  Call draw_item0(item(i%), conclusion)
 Next i%
 If para(1) = "0" Then
   If item(0).sig = "~" And para(0) = "1" Then
    If item(0).poi(1) = -6 Then
     Call is_three_angle_value(item(0).poi(0), 0, 0, "1", "0", "0", "", "", 0, 0, 0, -2000, 0, 0, 0, 0, 0, _
         0, 0, con_angle3_value(last_conclusion).data(0).data0, temp_record.record_data.data0.condition_data, 0)
               conclusion_data(last_conclusion).ty = angle3_value_
         GoTo draw_picture38_49_mark0
    ElseIf item(0).poi(1) = -10 Then '向量
       con_V_line_value(last_conclusion).data(0).v_line = item(0).poi(0)
        conclusion_data(last_conclusion).ty = V_line_value_
    End If
    If item(0).poi(0) > 0 And item(0).poi(1) > 0 Then
    Call is_line_value(item(0).poi(0), item(0).poi(1), 0, 0, 0, "", _
          0, -2000, 0, 0, 0, _
           con_line_value(last_conclusion).data(0).data0)
               conclusion_data(last_conclusion).ty = line_value_
         GoTo draw_picture38_49_mark0
    End If
   ElseIf item(0).sig = "/" Then
     If item(0).poi(0) > 0 And item(0).poi(1) > 0 Then
     con_relation(last_conclusion).data(1).poi(0) = item(0).poi(0)
      con_relation(last_conclusion).data(1).poi(1) = item(0).poi(1)
       con_relation(last_conclusion).data(1).poi(2) = item(0).poi(2)
        con_relation(last_conclusion).data(1).poi(3) = item(0).poi(3)
     con_relation(last_conclusion).data(1).n(0) = item(0).n(0)
      con_relation(last_conclusion).data(1).n(1) = item(0).n(1)
       con_relation(last_conclusion).data(1).n(2) = item(0).n(2)
        con_relation(last_conclusion).data(1).n(3) = item(0).n(3)
     con_relation(last_conclusion).data(1).line_no(0) = item(0).line_no(0)
      con_relation(last_conclusion).data(1).line_no(1) = item(0).line_no(1)
      Call is_relation(item(0).poi(0), item(0).poi(1), item(0).poi(2), _
       item(0).poi(3), item(0).n(0), item(0).n(1), item(0).n(2), _
        item(0).n(3), item(0).line_no(0), item(0).line_no(1), "", 0, -2000, _
          0, 0, 0, con_relation(last_conclusion).data(0), 0, 0, _
           conclusion_data(last_conclusion).ty, record_0.data0.condition_data, 1)
            Call con_relation_(last_conclusion, con_relation(last_conclusion).data(0))
             GoTo draw_picture38_49_mark0
    End If
   End If
 ElseIf para(2) = "0" Then
    If item(0).sig = "*" And item(1).sig = "*" And _
          item(0).poi(1) > 0 And item(1).poi(1) > 0 Then
      If regist_data.run_type = 1 And item(0).line_no(0) = item(0).line_no(1) And _
            item(1).line_no(0) = item(2).line_no(1) Then
            item(0).poi(0) = vector_number(item(0).poi(0), item(0).poi(1), m(0))
            item(0).poi(2) = vector_number(item(0).poi(2), item(0).poi(3), m(1))
            item(1).poi(0) = vector_number(item(1).poi(0), item(1).poi(1), m(2))
            item(1).poi(2) = vector_number(item(1).poi(2), item(1).poi(3), m(3))
            item(0).poi(1) = -10
            item(0).poi(3) = -10
            item(1).poi(1) = -10
            item(1).poi(3) = -10
         con_string = set_display_string_of_V_line(item(0).poi(0), False) + "*" + _
                    set_display_string_of_V_line(item(0).poi(2), False) + "-" + _
                    set_display_string_of_V_line(item(1).poi(0), False) + "*" + _
                    set_display_string_of_V_line(item(1).poi(2), False)
      Else
      If para(0) = time_string("-1", para(1), True, False) And is_zero Then
       con_dpoint_pair(last_conclusion).data(1).poi(0) = item(0).poi(0)
       con_dpoint_pair(last_conclusion).data(1).poi(1) = item(0).poi(1)
       con_dpoint_pair(last_conclusion).data(1).poi(2) = item(1).poi(0)
       con_dpoint_pair(last_conclusion).data(1).poi(3) = item(1).poi(1)
       con_dpoint_pair(last_conclusion).data(1).poi(4) = item(1).poi(2)
       con_dpoint_pair(last_conclusion).data(1).poi(5) = item(1).poi(3)
       con_dpoint_pair(last_conclusion).data(1).poi(6) = item(0).poi(2)
       con_dpoint_pair(last_conclusion).data(1).poi(7) = item(0).poi(3)
       conclusion_data(last_conclusion).ty = dpoint_pair_
     record_0.data0.condition_data.condition_no = 0 ' record0
     Call is_point_pair(item(0).poi(0), item(0).poi(1), item(1).poi(0), _
           item(1).poi(1), item(1).poi(2), item(1).poi(3), _
            item(0).poi(2), item(0).poi(3), item(0).n(0), item(0).n(1), _
             item(1).n(0), item(1).n(1), item(1).n(2), item(1).n(3), _
              item(0).n(2), item(0).n(3), item(0).line_no(0), item(1).line_no(0), _
               item(1).line_no(1), item(0).line_no(1), 0, -2000, 0, 0, 0, 0, _
             0, con_dpoint_pair(last_conclusion).data(0), _
               0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", record_0)
         GoTo draw_picture38_49_mark0
      End If
      End If
    ElseIf item(0).sig = "/" And item(1).sig = "/" Then
      If item(0).poi(0) > 0 And item(0).poi(1) > 0 And item(1).poi(0) > 0 And _
        item(1).poi(1) > 0 Then
         If para(0) = time_string("-1", para(1), True, False) And is_zero Then
           con_dpoint_pair(last_conclusion).data(1).poi(0) = item(0).poi(0)
           con_dpoint_pair(last_conclusion).data(1).poi(1) = item(0).poi(1)
           con_dpoint_pair(last_conclusion).data(1).poi(2) = item(0).poi(2)
           con_dpoint_pair(last_conclusion).data(1).poi(3) = item(0).poi(3)
           con_dpoint_pair(last_conclusion).data(1).poi(4) = item(2).poi(0)
           con_dpoint_pair(last_conclusion).data(1).poi(5) = item(2).poi(1)
           con_dpoint_pair(last_conclusion).data(1).poi(6) = item(2).poi(2)
           con_dpoint_pair(last_conclusion).data(1).poi(7) = item(2).poi(3)
           conclusion_data(last_conclusion).ty = dpoint_pair_
     record_0.data0.condition_data.condition_no = 0 ' record0
     Call is_point_pair(item(0).poi(0), item(0).poi(1), item(0).poi(2), _
           item(0).poi(3), item(1).poi(0), item(1).poi(1), _
            item(1).poi(2), item(1).poi(3), item(0).n(0), item(0).n(1), _
             item(0).n(2), item(0).n(3), item(1).n(0), item(1).n(1), _
              item(1).n(2), item(1).n(3), item(0).line_no(0), item(0).line_no(1), _
               item(1).line_no(0), item(1).line_no(1), 0, -2000, 0, 0, 0, 0, _
             0, con_dpoint_pair(last_conclusion).data(0), _
               0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", record_0)
                 GoTo draw_picture38_49_mark0
         End If
       ElseIf item(0).poi(0) = 0 And item(0).poi(1) = 0 And item(1).poi(0) = 0 And _
        item(1).poi(1) = 0 Then
         If para(0) = time_string("-1", para(1), True, False) And is_zero And _
             item(0).poi(3) > 0 And item(1).poi(3) Then
          con_eline(last_conclusion).data(1).data0.poi(0) = item(0).poi(2)
          con_eline(last_conclusion).data(1).data0.poi(1) = item(0).poi(3)
          con_eline(last_conclusion).data(1).data0.poi(2) = item(1).poi(2)
          con_eline(last_conclusion).data(1).data0.poi(3) = item(1).poi(3)
    Call is_equal_dline(item(0).poi(2), item(0).poi(3), item(1).poi(2), _
           item(1).poi(3), item(0).n(2), item(0).n(3), item(1).n(2), _
            item(1).n(3), item(0).line_no(1), item(1).line_no(1), 0, -2000, _
             0, 0, 0, con_eline(last_conclusion).data(0).data0, _
              0, 0, 0, "", record_0.data0.condition_data)
           conclusion_data(last_conclusion).ty = eline_
            GoTo draw_picture38_49_mark0
         ElseIf is_zero Then
     con_relation(last_conclusion).data(1).poi(0) = item(0).poi(2)
      con_relation(last_conclusion).data(1).poi(1) = item(0).poi(3)
       con_relation(last_conclusion).data(1).poi(2) = item(1).poi(2)
        con_relation(last_conclusion).data(1).poi(3) = item(1).poi(3)
     con_relation(last_conclusion).data(1).n(0) = item(0).n(2)
      con_relation(last_conclusion).data(1).n(1) = item(0).n(3)
       con_relation(last_conclusion).data(1).n(2) = item(1).n(2)
        con_relation(last_conclusion).data(1).n(3) = item(1).n(3)
     con_relation(last_conclusion).data(1).line_no(0) = item(0).line_no(1)
       con_relation(last_conclusion).data(1).line_no(1) = item(1).line_no(1)
          tv = divide_string(time_string("-1", para(0), False, False), _
               para(1), True, False)
         con_relation(last_conclusion).data(1).value = tv
          Call is_relation(item(0).poi(2), item(0).poi(3), _
           item(1).poi(2), item(1).poi(3), item(0).n(2), _
            item(0).n(3), item(1).n(2), item(1).n(3), item(0).line_no(1), _
             item(1).line_no(1), tv, 0, -2000, 0, 0, 0, _
                con_relation(last_conclusion).data(0), 0, 0, conclusion_data(last_conclusion).ty, _
                  record_0.data0.condition_data, 1)
            Call con_relation_(last_conclusion, con_relation(last_conclusion).data(0))
                          GoTo draw_picture38_49_mark0
         End If
      End If
     ElseIf item(0).sig = "~" And item(1).sig = "~" Then
      If para(0) = time_string("-1", para(1), True, False) And is_zero And _
          item(0).poi(1) > 0 And item(1).poi(1) > 0 Then
       con_eline(last_conclusion).data(1).data0.poi(0) = item(0).poi(0)
       con_eline(last_conclusion).data(1).data0.poi(1) = item(0).poi(1)
       con_eline(last_conclusion).data(1).data0.poi(2) = item(1).poi(0)
       con_eline(last_conclusion).data(1).data0.poi(3) = item(1).poi(1)
       conclusion_data(last_conclusion).ty = eline_
     Call is_equal_dline(item(0).poi(0), item(0).poi(1), item(1).poi(0), _
           item(1).poi(1), item(0).n(0), item(0).n(1), item(1).n(0), _
            item(1).n(1), item(0).line_no(0), item(1).line_no(0), 0, -2000, 0, 0, 0, _
              con_eline(last_conclusion).data(0).data0, _
             0, 0, 0, "", record_0.data0.condition_data)
        GoTo draw_picture38_49_mark0
      ElseIf is_zero Then
       con_relation(last_conclusion).data(1).poi(0) = item(0).poi(0)
       con_relation(last_conclusion).data(1).poi(1) = item(0).poi(1)
       con_relation(last_conclusion).data(1).poi(2) = item(1).poi(0)
       con_relation(last_conclusion).data(1).poi(3) = item(1).poi(1)
       con_relation(last_conclusion).data(1).n(0) = item(0).n(0)
       con_relation(last_conclusion).data(1).n(1) = item(0).n(1)
       con_relation(last_conclusion).data(1).n(2) = item(1).n(0)
       con_relation(last_conclusion).data(1).n(3) = item(1).n(1)
       con_relation(last_conclusion).data(1).line_no(0) = item(0).line_no(0)
       con_relation(last_conclusion).data(1).line_no(1) = item(1).line_no(0)
       Call is_relation(item(0).poi(0), item(0).poi(1), item(1).poi(0), _
        item(1).poi(1), item(0).n(0), item(0).n(1), item(1).n(0), _
         item(1).n(1), item(0).line_no(0), item(1).line_no(0), divide_string( _
          time_string("-1", para(1), False, False), para(0), True, False), 0, -2000, _
           0, 0, 0, con_relation(last_conclusion).data(0), _
            0, 0, conclusion_data(last_conclusion).ty, record_0.data0.condition_data, 1)
      Call con_relation_(last_conclusion, con_relation(last_conclusion).data(0))
               GoTo draw_picture38_49_mark0
       Else
       Call is_two_line_value(item(0).poi(0), item(0).poi(1), item(1).poi(0), _
        item(1).poi(1), item(0).n(0), item(0).n(1), item(1).n(0), _
         item(1).n(1), item(0).line_no(0), item(1).line_no(0), para(0), para(1), "", _
             0, -2000, 0, 0, 0, con_two_line_value(last_conclusion).data(0), 0, _
               record_0.data0.condition_data)
                    conclusion_data(last_conclusion).ty = two_line_value_
                     GoTo draw_picture38_49_mark0
    End If
    ElseIf item(0).sig = "/" And item(1).sig = "F" Then
     If item(0).poi(0) > 0 And item(0).poi(2) > 0 Then
       con_relation(last_conclusion).data(1).poi(0) = item(0).poi(0)
       con_relation(last_conclusion).data(1).poi(1) = item(0).poi(1)
       con_relation(last_conclusion).data(1).poi(2) = item(0).poi(2)
       con_relation(last_conclusion).data(1).poi(3) = item(0).poi(3)
       con_relation(last_conclusion).data(1).n(0) = item(0).n(0)
       con_relation(last_conclusion).data(1).n(1) = item(0).n(1)
       con_relation(last_conclusion).data(1).n(2) = item(0).n(2)
       con_relation(last_conclusion).data(1).n(3) = item(0).n(3)
       con_relation(last_conclusion).data(1).line_no(0) = item(0).line_no(0)
       con_relation(last_conclusion).data(1).line_no(1) = item(0).line_no(1)
       'con_relation(last_conclusion).data(1).value=
      Call is_relation(item(0).poi(0), item(0).poi(1), item(0).poi(2), item(0).poi(3), _
           item(0).n(0), item(0).n(1), item(0).n(2), item(0).n(3), item(0).line_no(0), _
            item(0).line_no(1), divide_string(time_string("-1", para(1), False, False), para(0), _
             True, False), 0, -2000, 0, 0, 0, con_relation(last_conclusion).data(0), _
              0, 0, 0, record_0.data0.condition_data, 1)
               conclusion_data(last_conclusion).ty = relation_
      Call con_relation_(last_conclusion, con_relation(last_conclusion).data(0))
               GoTo draw_picture38_49_mark0
      End If
    ElseIf item(0).sig = "F" And item(1).sig = "/" Then
     If item(1).poi(0) > 0 And item(1).poi(2) > 0 Then
       con_relation(last_conclusion).data(1).poi(0) = item(1).poi(0)
       con_relation(last_conclusion).data(1).poi(1) = item(1).poi(1)
       con_relation(last_conclusion).data(1).poi(2) = item(1).poi(2)
       con_relation(last_conclusion).data(1).poi(3) = item(1).poi(3)
       con_relation(last_conclusion).data(1).n(0) = item(1).n(0)
       con_relation(last_conclusion).data(1).n(1) = item(1).n(1)
       con_relation(last_conclusion).data(1).n(2) = item(1).n(2)
       con_relation(last_conclusion).data(1).n(3) = item(1).n(3)
       con_relation(last_conclusion).data(1).line_no(0) = item(1).line_no(0)
       con_relation(last_conclusion).data(1).line_no(1) = item(0).line_no(1)
      Call is_relation(item(1).poi(0), item(1).poi(1), item(1).poi(2), item(1).poi(3), _
           item(1).n(0), item(1).n(1), item(1).n(2), item(1).n(3), item(1).line_no(0), _
            item(1).line_no(1), divide_string(time_string("-1", para(0), False, False), para(1), _
             True, False), 0, -2000, 0, 0, 0, con_relation(last_conclusion).data(0), _
              0, 0, 0, record_0.data0.condition_data, 1)
               conclusion_data(last_conclusion).ty = relation_
      Call con_relation_(last_conclusion, con_relation(last_conclusion).data(0))
               GoTo draw_picture38_49_mark0
     End If
    ElseIf item(0).sig = "~" And item(1).sig = "F" Then
     If para(0) = "1" And item(0).poi(1) > 0 Then
     Call is_line_value(item(0).poi(0), item(0).poi(1), _
           0, 0, 0, "", 0, -2000, 0, 0, 0, _
        con_line_value(last_conclusion).data(0).data0)
      con_line_value(last_conclusion).data(0).data0.value = time_string("-1", _
         para(1), True, False)
      conclusion_data(last_conclusion).ty = line_value_
         GoTo draw_picture38_49_mark0
     End If
    ElseIf item(0).sig = "F" And item(1).sig = "~" Then
     If para(1) = "1" And item(1).poi(1) > 0 Then
     Call is_line_value(item(1).poi(0), item(1).poi(1), _
           0, 0, 0, "", 0, -2000, 0, 0, 0, _
        con_line_value(last_conclusion).data(0).data0)
      con_line_value(last_conclusion).data(0).data0.value = time_string( _
        "-1", para(0), True, False)
      conclusion_data(last_conclusion).ty = line_value_
         GoTo draw_picture38_49_mark0
    End If
   End If
 ElseIf para(3) = "0" Then
     'If item(0).sig = "~" And item(1).sig = "~" And item(2).sig = "~" Then
'set_con_general_string
  If item(0).sig = "~" And item(1).sig = "~" And item(2).sig = "F" And _
         item(0).poi(1) > 0 And item(1).poi(1) > 0 Then
      Call is_two_line_value( _
       item(0).poi(0), item(0).poi(1), item(1).poi(0), item(1).poi(1), _
        item(0).n(0), item(0).n(1), item(1).n(0), item(1).n(1), _
         item(0).line_no(0), item(1).line_no(0), para(0), para(1), _
          time_string("-1", para(2), True, False), 0, -2000, 0, 0, 0, _
           con_two_line_value(last_conclusion).data(0), 0, record_0.data0.condition_data)
     conclusion_data(last_conclusion).ty = two_line_value_
         GoTo draw_picture38_49_mark0
     ElseIf item(0).sig = "~" And item(1).sig = "F" And item(2).sig = "~" And _
            item(0).poi(1) > 0 And item(2).poi(1) > 0 Then
      Call is_two_line_value( _
       item(0).poi(0), item(0).poi(1), item(2).poi(0), item(2).poi(1), _
        item(0).n(0), item(0).n(1), item(2).n(0), item(2).n(1), _
         item(0).line_no(0), item(2).line_no(0), para(0), para(2), _
          time_string("-1", para(1), True, False), 0, -2000, 0, 0, 0, _
           con_two_line_value(last_conclusion).data(0), 0, record_0.data0.condition_data)
      conclusion_data(last_conclusion).ty = two_line_value_
         GoTo draw_picture38_49_mark0
     ElseIf item(0).sig = "F" And item(1).sig = "~" And item(2).sig = "~" And _
             item(1).poi(1) > 0 And item(2).poi(1) > 0 Then
      Call is_two_line_value( _
       item(1).poi(0), item(1).poi(1), item(2).poi(0), item(2).poi(1), _
        item(1).n(0), item(1).n(1), item(2).n(0), item(2).n(1), _
         item(1).line_no(0), item(2).line_no(0), para(1), para(2), _
          time_string("-1", para(2), True, False), 0, -2000, 0, 0, 0, _
           con_two_line_value(last_conclusion).data(0), 0, record_0.data0.condition_data)
      conclusion_data(last_conclusion).ty = two_line_value_
         GoTo draw_picture38_49_mark0
     ElseIf item(0).sig = "/" And item(1).sig = "/" And item(2).sig = "/" Then
      '1/AB+1/CD-2/EF
      If item(0).poi(0) = 0 And item(0).poi(1) = 0 And item(1).poi(0) = 0 And _
       item(1).poi(1) = 0 And item(2).poi(0) = 0 And item(2).poi(1) = 0 Then
        item(0) = set_item0_(item(2).poi(2), item(2).poi(3), _
           item(0).poi(2), item(0).poi(3), "/", _
            item(2).n(2), item(2).n(3), item(0).n(2), _
             item(0).n(3), item(2).line_no(1), item(0).line_no(1), "", "", "", 0, _
                "", temp_record.record_data.data0.condition_data)
        item(1) = set_item0_(item(2).poi(2), item(2).poi(3), _
           item(1).poi(2), item(1).poi(3), "/", _
            item(2).n(2), item(2).n(3), item(1).n(2), item(1).n(3), _
             item(2).line_no(1), item(1).line_no(1), "", "", "", 0, "", condition_data0)
        item(2) = set_item0_(0, 0, 0, 0, "", 0, 0, 0, 0, 0, 0, "", "", "", 0, _
                   "", temp_record.record_data.data0.condition_data)
     End If
  End If
 Else
  If item(0).sig = "/" And item(1).sig = "/" And item(2).sig = "/" And _
       item(3).sig = "/" And set_tv = "0" Then
   If item(0).poi(0) = 0 And item(0).poi(1) = 0 And item(1).poi(0) = 0 And _
       item(1).poi(1) = 0 And item(2).poi(0) = 0 And item(2).poi(1) = 0 And _
        item(3).poi(0) = 0 And item(3).poi(3) = 0 Then
   item(0) = set_item0_(item(3).poi(2), item(3).poi(3), _
           item(0).poi(2), item(0).poi(3), "/", _
            item(3).n(2), item(3).n(3), item(0).n(2), item(0).n(3), _
             item(3).line_no(1), item(0).line_no(1), "", "", "", 0, _
                "", temp_record.record_data.data0.condition_data)
   item(1) = set_item0_(item(3).poi(2), item(3).poi(3), _
           item(1).poi(2), item(1).poi(3), "/", _
             item(3).n(2), item(3).n(3), item(1).n(2), _
              item(1).n(3), item(3).line_no(1), item(1).line_no(1), "", "", "", 0, _
                "", temp_record.record_data.data0.condition_data)
   item(2) = set_item0_(item(3).poi(2), item(3).poi(3), _
           item(2).poi(2), item(2).poi(3), "/", _
           item(3).n(2), item(3).n(3), item(2).n(2), item(2).n(3), _
            item(3).line_no(1), item(2).line_no(1), "", "", "", 0, "", condition_data0)
   item(3) = set_item0_(0, 0, 0, 0, "", 0, 0, 0, 0, 0, 0, "", "", "", 0, _
                 "", temp_record.record_data.data0.condition_data)
   End If
  'ElseIf item(0).sig = "~" And item(1).sig = "~" And item(2).sig = "~" Then
  ' it_(0) = item_number(set_item0_(item(0).poi(0), item(0).poi(1), _
           item(3).poi(0), item(3).poi(1), "/", _
            item(0).n(0), item(0).n(1), item(3).n(0), item(3).n(1), _
              item(0).line_no(0), item(3).line_no(0), "", "", "","" 0, condition_data0))
 '  it_(1) = item_number(set_item0_(item(1).poi(0), item(1).poi(1), _
           item(3).poi(0), item(3).poi(1), "/", _
            item(1).n(0), item(1).n(1), item(3).n(0), item(3).n(1), _
             item(1).line_no(0), item(3).line_no(0), "", "", "","" 0, condition_data0))
 '  it_(2) = item_number(set_item0_(item(2).poi(0), item(2).poi(1), _
           item(3).poi(0), item(3).poi(1), "/", _
            item(2).n(0), item(2).n(1), item(3).n(0), item(3).n(1), _
             item(2).line_no(0), item(3).line_no(0), "", "", "","" 0, condition_data0))
 '  it_(3) = 0 ' item_number(set_item0_(0, 0, 0, 0, "", 0, 0, 0, 0, 0, 0, "", "", "",  0, condition_data0))
 '  para_(0) = para(0)
 '  para_(1) = para(1)
 '  para_(2) = para(2)
 '  para_(3) = para(3)
 '  tv_ = tv
 End If
 End If
 End If
   it(0) = item_number(item(0), it_para(0), temp_record.record_data.data0.condition_data)
   it(1) = item_number(item(1), it_para(1), temp_record.record_data.data0.condition_data)
   it(2) = item_number(item(2), it_para(2), temp_record.record_data.data0.condition_data)
   para(0) = time_string(para(0), it_para(0), True, False)
   para(1) = time_string(para(1), it_para(1), True, False)
   para(2) = time_string(para(2), it_para(2), True, False)
    If set_tv <> "0" And set_tv <> "" Then
     If it(3) = 0 Then
      para(3) = add_string(para(3), set_tv, True, False)
     Else
      para(3) = add_string(set_tv, set_new_value(item(), para(), 3, 3), True, False)
     End If
     it(3) = 0
   Else
    it(3) = item_number(item(3), it_para(3), temp_record.record_data.data0.condition_data)
    para(3) = time_string(para(3), it_para(3), True, False)
   End If
    For i% = 0 To 3
    Call draw_item0(item0(it(i%)).data(0), conclusion)
    Next i%
temp_record.record_.conclusion_no = last_conclusion + 1
If temp_record.record_data.data0.condition_data.condition_no = 0 Then
 temp_record.record_data.data0.condition_data.condition_no = 254
End If
 If C_display_wenti.m_no(num) = 73 Then
   temp_record.record_.conclusion_ty = 73 '定值
 ElseIf C_display_wenti.m_no(num) = 75 Then
   temp_record.record_.conclusion_ty = 75 '极大值
 ElseIf C_display_wenti.m_no(num) = 76 Then
   temp_record.record_.conclusion_ty = 76 '极小值
 End If

'If c_display_wenti.m_no = 38 Then
 'tv = "0"
'Else
 tv = ""
'End If
'Call is_general_string(it(0), it(1), it(2), it(3), para(0), para(1), _
           para(2), para(3), tv, 0, -2000, 0, 0, 0, _
            con_general_string(last_conclusion).data(0), last_conclusion + 1, 0, _
              temp_record.record_data, 255)
 If C_display_wenti.m_no(num) = 38 Then
 con_general_string(last_conclusion).data(0).value = "0"
 Else
 con_general_string(last_conclusion).data(0).value = ""
 End If
temp_record.record_.display_no = 0
con_general_string(last_conclusion).data(0).item(0) = it(0)
con_general_string(last_conclusion).data(0).item(1) = it(1)
con_general_string(last_conclusion).data(0).item(2) = it(2)
con_general_string(last_conclusion).data(0).item(3) = it(3)
con_general_string(last_conclusion).data(0).para(0) = para(0)
con_general_string(last_conclusion).data(0).para(1) = para(1)
con_general_string(last_conclusion).data(0).para(2) = para(2)
con_general_string(last_conclusion).data(0).para(3) = para(3)
con_general_string(last_conclusion).data(0).value = "0"
'  con_general_string(last_conclusion).data(0).item(1), con_general_string(last_conclusion).data(0).item(2), _
'  con_general_string(last_conclusion).data(0).item(3), con_general_string(last_conclusion).data(0).para(0), _
'  con_general_string(last_conclusion).data(0).para(1), con_general_string(last_conclusion).data(0).para(2), _
'  con_general_string(last_conclusion).data(0).para(3), "", last_conclusion + 1, 0, 4, temp_record, 0, 0)
no% = 0
Call set_general_string(it(0), it(1), it(2), it(3), _
    para(0), para(1), para(2), para(3), "", last_conclusion + 1, 0, 4, temp_record, no%, 0)
    general_string(no%).display_con_string = con_string
  conclusion_data(last_conclusion).ty = general_string_
   ge_reduce_level = 3
End If
draw_picture38_49_mark0:
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
draw_picture38_49_mark1:
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
End Sub


Public Function draw_free_point(p%, c As String) As Boolean
Dim ele1 As condition_type
Dim ele2 As condition_type
'Dim temp_point_ As temp_element
If m_poi(p%).data(0).data0.coordinate.X <> 10000 Or m_poi(p%).data(0).data0.coordinate.Y <> 10000 Or _
   event_statue = wait_for_modify_char Then
 Exit Function '已经输入
End If
If p% > 0 Then
 Exit Function
Else
   If Asc(c) > 64 And _
           Asc(c) < 91 Then
    p% = point_number(c)
         '读出条件的点, 记录点号
    If p% > 0 Then
     Exit Function
    End If
   Else
    Exit Function
   End If

draw_free_point0:
 event_statue = wait_for_draw_point  '输点状态
   While event_statue = wait_for_draw_point '等待事件发生
    DoEvents
   Wend
 'If event_statue = wait_for_modify_char Then
  
  'Exit Function
If event_statue = draw_point_down Or event_statue = _
             draw_point_move Or event_statue = _
                    draw_point_up Then 'mouse_type <> 1 Then
    t_coord = input_coord
    'temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Or _
      event_statue = wait_for_modify_char Then
  draw_free_point = True
   Exit Function
 Else
     GoTo draw_free_point0
 End If
  input_point_type% = read_inter_point(t_coord, ele1, _
                              ele2, temp_point(0).no, True, 0)
          Call set_point_no_reduce(temp_point(0).no, 0)
     If input_point_type% <> new_free_point Then  '不是新的自由点
         If input_point_type% <> exist_point Then  '不是旧的自由点
          Call remove_point(temp_point(0).no, display, 0) '抹掉
         Else
          GoTo draw_free_point0
         End If
     End If
      Call set_point_name(temp_point(0).no, c) '
         p% = temp_point(0).no
          'Call put_name(p%)
    End If
End Function

Public Sub draw_picture24(ByVal num As Integer)
'□、□、□三点共线
Dim i%, tn%
Dim c_data0 As condition_data_type
Dim tp(2) As Integer
For i% = 0 To 2
  If m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree > 2 And i% <> 3 Then
    m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree = _
     m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree - 3
   End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
If compare_two_point(m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate, _
                        m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate, _
                         C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 2), 4) = _
   compare_two_point(m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate, _
                        m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate, _
                         C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 2), 4) Then
                         tp(0) = C_display_wenti.m_point_no(num, 0)
                         tp(1) = C_display_wenti.m_point_no(num, 1)
                         tp(2) = C_display_wenti.m_point_no(num, 2)
ElseIf compare_two_point(m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate, _
                        m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate, _
                         C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 0), 4) = _
   compare_two_point(m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate, _
                        m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate, _
                         C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 0), 4) Then
                         tp(0) = C_display_wenti.m_point_no(num, 1)
                         tp(1) = C_display_wenti.m_point_no(num, 2)
                         tp(2) = C_display_wenti.m_point_no(num, 0)
ElseIf compare_two_point(m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate, _
                        m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate, _
                         C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 1), 4) = _
   compare_two_point(m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate, _
                        m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate, _
                         C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 1), 4) Then
                         tp(0) = C_display_wenti.m_point_no(num, 2)
                         tp(1) = C_display_wenti.m_point_no(num, 0)
                         tp(2) = C_display_wenti.m_point_no(num, 1)
End If

Call line_number(tp(0), tp(1), _
                     pointapi0, pointapi0, _
                     depend_condition(0, 0), _
                     depend_condition(0, 0), _
                     conclusion, conclusion_color, 1, 0)
Call line_number(tp(1), tp(2), _
                     pointapi0, pointapi0, _
                     depend_condition(0, 0), _
                     depend_condition(0, 0), _
                     conclusion, conclusion_color, 1, 0)
'Call add_point_to_line(C_display_wenti.m_point_no(num, 2), _
           con_l%, 0, True, True, 0,  c_data0)
 conclusion_data(last_conclusion).ty = point3_on_line_
  con_Three_point_on_line(last_conclusion).data(0).poi(0) = C_display_wenti.m_point_no(num, 0)
    con_Three_point_on_line(last_conclusion).data(0).poi(1) = C_display_wenti.m_point_no(num, 1)
     con_Three_point_on_line(last_conclusion).data(0).poi(2) = C_display_wenti.m_point_no(num, 2)
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
End Sub
'Public Sub draw_picture22(ByVal num As Integer, no_reduce As Byte)
'Dim i%
'For i% = 0 To 3
'If draw_free_point(c_display_wenti.m_point_no(num,i%), _
      c_display_wenti.m_condition(i%)) Then
'End If
 ' If poi(c_display_wenti.m_point_no(num,i%)).data(0).degree > 2 And i% <> 3 Then
  '  poi(c_display_wenti.m_point_no(num,i%)).data(0).degree = _
   '  poi(c_display_wenti.m_point_no(num,i%)).data(0).degree - 3
   'End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
'Next i%
'End Sub

Public Sub draw_picture23(ByVal num As Integer)
'□、□、□、□四点共圆
Dim i%, tn%
Dim temp_record As total_record_type
Dim c_data0 As circle_data_type
For i% = 0 To 3
  If m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree > 2 And i% <> 3 Then
    m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree = _
     m_poi(C_display_wenti.m_point_no(num, i%)).data(0).degree - 3
   End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
      tn% = m_circle_number(1, 0, pointapi0, _
             C_display_wenti.m_point_no(num, 0), _
              C_display_wenti.m_point_no(num, 1), _
               C_display_wenti.m_point_no(num, 2), _
                0, 0, 0, 1, 1, conclusion, conclusion_color, True)
      Call C_display_wenti.set_m_point_no(0, tn%, 12, False)
If event_statue = input_prove_by_hand Then
temp_record.record_.display_no = -(num + 1)
  Call set_four_point_on_circle(C_display_wenti.m_point_no(num, 0), _
   C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
    C_display_wenti.m_point_no(num, 3), 0, temp_record, 0, 0)
Else
 If C_display_wenti.m_point_no(num, 12) = 0 Then
     Call C_display_wenti.set_m_point_no(num, tn%, 12, False)
 End If
 conclusion_data(last_conclusion).ty = point4_on_circle_
 If is_four_point_on_circle(C_display_wenti.m_point_no(num, 0), _
      C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
       C_display_wenti.m_point_no(num, 3), tn%, _
            con_Four_point_on_circle(last_conclusion).data(0), False) Then
     conclusion_data(last_conclusion).no(0) = tn%
 End If
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
End If
End Sub

Public Sub draw_picture26(ByVal num As Integer)
'26 点□是线段□□的中点
Dim i%, tn%
Dim con_ty As Byte
Dim temp_record As total_record_type
For i% = 0 To 2
  If i% <> 3 Then
     Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
If set_or_prove < 2 Then
 Call line_number(C_display_wenti.m_point_no(num, 1), _
                  C_display_wenti.m_point_no(num, 2), _
                  pointapi0, pointapi0, _
                  depend_condition(0, 0), depend_condition(0, 0), _
                  conclusion, condition_color, 1, 0)
'End If
con_eline(last_conclusion).data(1).data0.poi(0) = C_display_wenti.m_point_no(num, 0)
con_eline(last_conclusion).data(1).data0.poi(1) = C_display_wenti.m_point_no(num, 1)
con_eline(last_conclusion).data(1).data0.poi(2) = C_display_wenti.m_point_no(num, 2)
con_eline(last_conclusion).data(1).data0.poi(3) = C_display_wenti.m_point_no(num, 3)
Call is_equal_dline(C_display_wenti.m_point_no(num, 1), _
    C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 0), _
       C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, 0, tn%, -2000, 0, 0, 0, _
         con_eline(last_conclusion).data(0).data0, 0, 0, con_ty, "", record_0.data0.condition_data)
'con_eline(last_conclusion).old_data = con_eline(last_conclusion).data
conclusion_data(last_conclusion).ty = con_ty
conclusion_data(last_conclusion).no(0) = tn%
If con_ty = midpoint_ Then
con_mid_point(last_conclusion).data(0).poi(0) = _
         con_eline(last_conclusion).data(0).data0.poi(0)
con_mid_point(last_conclusion).data(0).poi(1) = _
         con_eline(last_conclusion).data(0).data0.poi(1)
con_mid_point(last_conclusion).data(0).poi(2) = _
         con_eline(last_conclusion).data(0).data0.poi(3)
End If
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
Call set_mid_point(C_display_wenti.m_point_no(num, 1), _
   C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 2), _
    0, 0, 0, 0, 0, temp_record, 0, 0, 0, 0, 3)
End If
End Sub

Public Sub draw_picture30(ByVal num As Integer)
'∠□□□=∠□□□
Dim ang(1) As Integer
Dim dn(2) As Integer
Dim con_ty As Byte
Dim temp_record As total_record_type
Dim i%
For i% = 0 To 6
  If i% <> 3 Then
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
If set_or_prove < 2 Then
ang(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 0), _
                          C_display_wenti.m_point_no(num, 1), _
                          C_display_wenti.m_point_no(num, 2), 0, 0))
 If ang(0) <> 0 Then
  Call draw_angle(C_display_wenti.m_point_no(num, 0), _
                  C_display_wenti.m_point_no(num, 1), _
                  C_display_wenti.m_point_no(num, 2), conclusion)
  ang(1) = Abs(angle_number(C_display_wenti.m_point_no(num, 3), _
                            C_display_wenti.m_point_no(num, 4), _
                            C_display_wenti.m_point_no(num, 5), 0, 0))
   If ang(1) <> 0 Then
     Call draw_angle(C_display_wenti.m_point_no(num, 3), _
                     C_display_wenti.m_point_no(num, 4), _
                     C_display_wenti.m_point_no(num, 5), conclusion)
conclusion_data(last_conclusion).ty = angle3_value_
If ang(0) < ang(1) Then
con_angle3_value(last_conclusion).data(0).data0.angle(0) = ang(0)
con_angle3_value(last_conclusion).data(0).data0.angle(1) = ang(1)
con_angle3_value(last_conclusion).data(0).data0.angle(2) = 0
Else
con_angle3_value(last_conclusion).data(0).data0.angle(0) = ang(1)
con_angle3_value(last_conclusion).data(0).data0.angle(1) = ang(0)
con_angle3_value(last_conclusion).data(0).data0.angle(2) = 0
End If
con_angle3_value(last_conclusion).data(0).data0.para(0) = "1"
con_angle3_value(last_conclusion).data(0).data0.para(1) = "-1"
con_angle3_value(last_conclusion).data(0).data0.para(2) = "0"
con_angle3_value(last_conclusion).data(0).data0.value = "0"
If is_equal_angle(ang(0), ang(1), dn(0), dn(1)) Then
 If dn(0) > 0 Then
       conclusion_data(last_conclusion).no(0) = dn(0)
 ElseIf con_ty = angle3_value_ Then
  Call add_conditions_to_record(angle3_value_, dn(0), dn(1), 0, temp_record.record_data.data0.condition_data)
If set_three_angle_value(con_angle3_value(last_conclusion).data(0).data0.angle(0), _
  con_angle3_value(last_conclusion).data(0).data0.angle(1), 0, "1", "-1", "0", _
   "0", 0, temp_record, conclusion_data(last_conclusion).no(0), 0, 0, 0, 0, 0, False) = False Then
                       conclusion_data(last_conclusion).no(0) = 0
End If
End If
End If
End If
End If
Else
record_0.data0.condition_data.condition_no = 0 'record0
Call set_three_angle_value(Abs(angle_number(C_display_wenti.m_point_no(num, 0), _
     C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), 0, 0)), _
       Abs(angle_number(C_display_wenti.m_point_no(num, 3), _
        C_display_wenti.m_point_no(num, 4), C_display_wenti.m_point_no(num, 5), 0, 0)), _
         0, "1", "-1", "0", "0", 0, temp_record, 0, 0, 0, 3, 0, 0, False)
End If
End Sub

Public Sub draw_picture32(ByVal num As Integer)
'32 □□/□□＝□□/□□
Dim i%, n%
Dim tl As Integer
Dim temp_record As total_record_type
For i% = 0 To 7
  If i% <> 3 Then
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
If set_or_prove < 2 Then
i% = line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                 conclusion, conclusion_color, 1, 0)
i% = line_number(C_display_wenti.m_point_no(num, 2), _
                 C_display_wenti.m_point_no(num, 3), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                 conclusion, conclusion_color, 1, 0)
i% = line_number(C_display_wenti.m_point_no(num, 4), _
                 C_display_wenti.m_point_no(num, 5), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                 conclusion, conclusion_color, 1, 0)
i% = line_number(C_display_wenti.m_point_no(num, 6), _
                 C_display_wenti.m_point_no(num, 7), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                 conclusion, conclusion_color, 1, 0)
For i% = 0 To 3
tl = line_number0(C_display_wenti.m_point_no(num, 2 * i%), _
    C_display_wenti.m_point_no(num, 2 * i% + 1), 0, 0)
con_dpoint_pair(last_conclusion).data(1).line_no(i%) = tl
con_dpoint_pair(last_conclusion).data(1).poi(2 * i%) = _
      C_display_wenti.m_point_no(num, 2 * i%)
con_dpoint_pair(last_conclusion).data(1).poi(2 * i% + 1) = _
      C_display_wenti.m_point_no(num, 2 * i% + 1)
Next i%
conclusion_data(last_conclusion).ty = dpoint_pair_
record_0.data0.condition_data.condition_no = 0 ' record0
Call is_point_pair(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
  C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
   C_display_wenti.m_point_no(num, 4), C_display_wenti.m_point_no(num, 5), _
    C_display_wenti.m_point_no(num, 6), C_display_wenti.m_point_no(num, 7), _
     0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -2000, 0, 0, 0, 0, _
      0, con_dpoint_pair(last_conclusion).data(0), _
       0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", record_0)
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
record_0.data0.condition_data.condition_no = 0 ' record0
Call set_dpoint_pair(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
  C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
   C_display_wenti.m_point_no(num, 4), C_display_wenti.m_point_no(num, 5), _
    C_display_wenti.m_point_no(num, 6), C_display_wenti.m_point_no(num, 7), _
     0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, temp_record, False, 0, 0, 0, 3, True)
End If
End Sub

Public Sub draw_picture35_54(ByVal num As Integer)
'35□□=?
Dim value1 As String
Dim i%
Dim temp_record As total_record_type
If C_display_wenti.m_no(num) = 54 Then
i% = 2
While Asc(C_display_wenti.m_condition(num, i%)) > 13 ' c_display_wenti.m_condition(num,i%) <> empty_char
 If C_display_wenti.m_condition(num, i%) < "A" Then
value1 = value1 + C_display_wenti.m_condition(num, i%)
i% = i% + 1
 Else
 value1 = value1 + C_display_wenti.m_condition(num, i%)
 i% = i% + 1
End If
Wend
End If

If set_or_prove < 2 Then
Call line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                 conclusion, conclusion_color, 1, 0)
If C_display_wenti.m_no(num) = 54 Then
 con_line_value(last_conclusion).data(0).data0.value = value1
End If
conclusion_data(last_conclusion).ty = line_value_
     Call is_line_value(C_display_wenti.m_point_no(num, 0), _
        C_display_wenti.m_point_no(num, 1), _
          0, 0, 0, "", 0, -2000, 0, 0, 0, _
        con_line_value(last_conclusion).data(0).data0)
conclusion_data(last_conclusion).wenti_no = num
 last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
record_0.data0.condition_data.condition_no = 0 'record0
Call set_line_value(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), value1, 0, 0, 0, temp_record.record_data, 0, 3, False)
 Exit Sub
End If
End Sub

Public Sub draw_picture36_53(ByVal num As Integer)
'36 ∠□□□=?
'53 ∠□□□=!_~°
Dim value1 As String
Dim temp_record As total_record_type
Dim i%, bra%
Dim tn(2) As Integer
If C_display_wenti.m_no(num) = 53 Then
i% = 3
While Asc(C_display_wenti.m_condition(num, i%)) > 13 ' c_display_wenti.m_condition(num,i%) <> empty_char
 If C_display_wenti.m_condition(num, i%) < "A" Then
value1 = value1 + C_display_wenti.m_condition(num, i%)
i% = i% + 1
 Else
 value1 = value1 + C_display_wenti.m_condition(num, i%)
 i% = i% + 1
End If
Wend
End If
If set_or_prove < 2 Then
Call line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                 conclusion, conclusion_color, 1, 0)
Call line_number(C_display_wenti.m_point_no(num, 2), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                 conclusion, conclusion_color, 1, 0)
If C_display_wenti.m_no(num) = 53 Then
con_angle3_value(last_conclusion).data(0).data0.value = value1
End If
'
con_angle3_value(last_conclusion).data(1).data0.angle(0) = _
 Abs(angle_number( _
   C_display_wenti.m_point_no(num, 0), _
    C_display_wenti.m_point_no(num, 1), _
      C_display_wenti.m_point_no(num, 2), 0, 0))
con_angle3_value(last_conclusion).data(1).data0.angle(1) = 0
con_angle3_value(last_conclusion).data(1).data0.angle(2) = 0
con_angle3_value(last_conclusion).data(1).data0.para(0) = "1"
con_angle3_value(last_conclusion).data(1).data0.para(1) = "0"
con_angle3_value(last_conclusion).data(1).data0.para(2) = "0"
con_angle3_value(last_conclusion).data(1).data0.value = ""
If is_three_angle_value(con_angle3_value(last_conclusion).data(1).data0.angle(0), _
      0, 0, "1", "0", "0", con_angle3_value(last_conclusion).data(1).data0.value, _
         "", conclusion_data(last_conclusion).no(0), 0, 0, 0, -1000, 0, 0, 0, 0, 0, 0, _
          con_angle3_value(last_conclusion).data(0).data0, _
             temp_record.record_data.data0.condition_data, 0) = False Then
      conclusion_data(last_conclusion).no(0) = 0
End If
            conclusion_data(last_conclusion).ty = angle3_value_
conclusion_data(last_conclusion).wenti_no = num
  last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
record_0.data0.condition_data.condition_no = 0 ' record0
Call set_angle_value(Abs(angle_number(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), 0, 0)), _
  value1, temp_record, 0, 0, False) '-1)
Exit Sub
End If
End Sub

Public Sub draw_picture37(ByVal num As Integer)
'□□/□□=?
Dim tn%, tn1%, tn2%, bra%
Dim cond_ty As Byte
Dim v As String
Dim tv As String
Dim temp_record As total_record_type
Call line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                 conclusion, conclusion_color, 1, 0)
Call line_number(C_display_wenti.m_point_no(num, 2), _
                 C_display_wenti.m_point_no(num, 3), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                 conclusion, conclusion_color, 1, 0)
       con_relation(last_conclusion).data(1).poi(0) = C_display_wenti.m_point_no(num, 0)
       con_relation(last_conclusion).data(1).poi(1) = C_display_wenti.m_point_no(num, 1)
       con_relation(last_conclusion).data(1).poi(2) = C_display_wenti.m_point_no(num, 2)
       con_relation(last_conclusion).data(1).poi(3) = C_display_wenti.m_point_no(num, 3)
con_relation(last_conclusion).data(1).poi(0) = C_display_wenti.m_point_no(num, 0)
con_relation(last_conclusion).data(1).poi(1) = C_display_wenti.m_point_no(num, 1)
con_relation(last_conclusion).data(1).poi(2) = C_display_wenti.m_point_no(num, 2)
con_relation(last_conclusion).data(1).poi(3) = C_display_wenti.m_point_no(num, 3)
If is_relation(C_display_wenti.m_point_no(num, 0), _
  C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
   C_display_wenti.m_point_no(num, 3), 0, 0, 0, 0, 0, 0, _
     con_relation(last_conclusion).data(1).value, tn%, _
      -2000, 0, 0, 0, con_relation(last_conclusion).data(0), _
       tn1%, tn2%, cond_ty, record_0.data0.condition_data, 1) Then
  If cond_ty <> line_value_ Then
   If con_relation(last_conclusion).data(0).poi(0) = C_display_wenti.m_point_no(num, 0) And _
       con_relation(last_conclusion).data(0).poi(1) = C_display_wenti.m_point_no(num, 1) And _
        con_relation(last_conclusion).data(0).poi(2) = C_display_wenti.m_point_no(num, 2) And _
         con_relation(last_conclusion).data(0).poi(3) = C_display_wenti.m_point_no(num, 3) Then
    conclusion_data(last_conclusion).ty = cond_ty
     conclusion_data(last_conclusion).no(0) = tn%
   Else
    temp_record.record_data.data0.condition_data.condition_no = 1
    temp_record.record_data.data0.condition_data.condition(1).ty = cond_ty
    temp_record.record_data.data0.condition_data.condition(1).no = tn%
    temp_record.record_data.data0.theorem_no = 1
    Call set_Drelation0(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
       C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
        con_relation(last_conclusion).data(1).value, tn%, cond_ty, _
         temp_record)
     conclusion_data(last_conclusion).ty = cond_ty
     conclusion_data(last_conclusion).no(0) = tn%
    End If
  Else
    temp_record.record_data.data0.condition_data.condition_no = 0
    Call add_conditions_to_record(cond_ty, tn1%, tn2%, 0, _
         temp_record.record_data.data0.condition_data)
    Call set_Drelation0(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 0), _
       C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 0), _
          con_relation(last_conclusion).data(1).value, tn%, cond_ty, _
           temp_record)
     conclusion_data(last_conclusion).ty = cond_ty
     conclusion_data(last_conclusion).no(0) = tn%
   End If
  Else
  conclusion_data(last_conclusion).ty = relation_
  End If
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True

End Sub

Public Sub draw_picture33_34(ByVal num As Integer)
'33△□□□∽△□□□
'34△□□□≌△□□□
Dim i%
Dim dir(1) As Integer
Dim temp_record As total_record_type
For i% = 0 To 3
  If i% <> 3 Then
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
If set_or_prove < 2 Then
Call draw_triangle(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   C_display_wenti.m_point_no(num, 2), conclusion)
Call draw_triangle(C_display_wenti.m_point_no(num, 3), _
                   C_display_wenti.m_point_no(num, 4), _
                   C_display_wenti.m_point_no(num, 5), conclusion)
If C_display_wenti.m_no(num) = 33 Then
conclusion_data(last_conclusion).ty = similar_triangle_
Call is_similar_triangle(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 4), _
   C_display_wenti.m_point_no(num, 5), 0, -1000, 0, 0, _
     con_similar_triangle(last_conclusion).data(0), record_0, 0)
Else
conclusion_data(last_conclusion).ty = total_equal_triangle_
Call is_total_equal_triangle1(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 4), _
   C_display_wenti.m_point_no(num, 5), 0, -1000, 0, 0, _
    con_total_equal_triangle(last_conclusion).data(0), record_0, 0)
End If
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
record_0.data0.condition_data.condition_no = 0 ' record0
If C_display_wenti.m_no(num) = 33 Then
Call set_similar_triangle(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 4), _
   C_display_wenti.m_point_no(num, 5), temp_record, 0, 3, 0)
ElseIf C_display_wenti.m_no(num) = 34 Then
Call set_total_equal_triangle(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 4), _
   C_display_wenti.m_point_no(num, 5), temp_record, 0, 0)
End If
End If
End Sub
Public Sub draw_picture68_69(ByVal num As Integer)
'68 △□□□与△□□□面积相等
'69 ∠□□□/∠□□□=!_~
Dim i%
Dim dir(1) As Integer
Dim ty As Byte
Dim area_ele(2) As condition_type
Dim temp_record As total_record_type
For i% = 0 To 3
  If i% <> 3 Then
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
If set_or_prove < 2 Then
Call draw_triangle(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   C_display_wenti.m_point_no(num, 2), conclusion)
Call draw_triangle(C_display_wenti.m_point_no(num, 3), _
                   C_display_wenti.m_point_no(num, 4), _
                   C_display_wenti.m_point_no(num, 5), conclusion)
'End If
If num = 68 Then
area_ele(0).ty = triangle_
area_ele(0).no = triangle_number(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, 0, 0)
area_ele(1).ty = triangle_
area_ele(1).no = triangle_number(C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 4), _
   C_display_wenti.m_point_no(num, 5), 0, 0, 0, 0, 0, 0, 0)
Call is_area_relation(area_ele(0), area_ele(1), "1", 0, 0, 0, 0, _
      con_area_relation(last_conclusion).data(0).area_element(0), _
       con_area_relation(last_conclusion).data(0).area_element(1), _
        con_area_relation(last_conclusion).data(0).area_element(2), _
          con_area_relation(last_conclusion).data(0).value, ty, 0, 0) 'Then
   For i% = 0 To 2
    If con_area_relation(last_conclusion).data(0).area_element(i).no > 0 Then
     If triangle(con_area_relation(last_conclusion).data(0).area_element(i).no). _
        data(0).poi(2) > 3 Then
         last_area_element_in_conclusion = last_area_element_in_conclusion + 1
           ReDim Preserve Area_element_in_conclusion(last_area_element_in_conclusion) As condition_type
            Area_element_in_conclusion(last_area_element_in_conclusion) = _
                 con_area_relation(last_conclusion).data(0).area_element(i)
     End If
    End If
   Next i%
 conclusion_data(last_conclusion).ty = area_relation_
'End If
Else
area_ele(0).ty = triangle_
area_ele(0).no = triangle_number(C_display_wenti.m_point_no(num, 0), _
                    C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, 0, 0)
area_ele(1).ty = triangle_
area_ele(1).no = triangle_number(C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 4), _
         C_display_wenti.m_point_no(num, 5), 0, 0, 0, 0, 0, 0, 0)
Call is_area_relation(area_ele(0), area_ele(1), "", 0, 0, 0, 0, _
    con_area_relation(last_conclusion).data(0).area_element(0), _
     con_area_relation(last_conclusion).data(0).area_element(1), _
      con_area_relation(last_conclusion).data(0).area_element(2), "", ty, 0, 0) 'Then
   For i% = 0 To 2
    If con_area_relation(last_conclusion).data(0).area_element(i).no > 0 Then
     If triangle(con_area_relation(last_conclusion).data(0).area_element(i%).no). _
        data(0).poi(3) > 3 Then
         last_area_element_in_conclusion = last_area_element_in_conclusion + 1
          ReDim Preserve Area_element_in_conclusion(last_area_element_in_conclusion) As condition_type
            Area_element_in_conclusion(last_area_element_in_conclusion) = _
                 con_area_relation(last_conclusion).data(0).area_element(i)
     End If
    End If
   Next i%
'Else
 conclusion_data(last_conclusion).ty = area_relation_
'End If

End If
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
th_chose(157).chose = 1
th_chose(158).chose = 1
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
End If
End Sub

Public Sub draw_picture31(ByVal num As Integer)
'□□/□□=!_~
Dim i%, j%, tn%
Dim value1 As String
Dim temp_record As total_record_type
'con_relation(last_conclusion).data(0).poi(0) = c_display_wenti.m_point_no(num,0)
'con_relation(last_conclusion).data(0).poi(1) = c_display_wenti.m_point_no(num,1)
'con_relation(last_conclusion).data(0).poi(2) = c_display_wenti.m_point_no(num,2)
'con_relation(last_conclusion).data(0).poi(3) = c_display_wenti.m_point_no(num,3)
For i% = 0 To 3
  If i% <> 3 Then
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
value1 = initial_string(number_string(C_display_wenti.m_point_no(num, 4))) 'initial_string(cond_to_string(num, 4, 18, 0))
If set_or_prove < 2 Then
Call line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                 conclusion, conclusion_color, 1, 0)
Call line_number(C_display_wenti.m_point_no(num, 2), _
                 C_display_wenti.m_point_no(num, 3), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                 conclusion, conclusion_color, 1, 0)
con_relation(last_conclusion).data(1).poi(0) = _
      C_display_wenti.m_point_no(num, 0)
con_relation(last_conclusion).data(1).poi(1) = _
      C_display_wenti.m_point_no(num, 1)
con_relation(last_conclusion).data(1).poi(2) = _
      C_display_wenti.m_point_no(num, 2)
con_relation(last_conclusion).data(1).poi(3) = _
      C_display_wenti.m_point_no(num, 3)
con_relation(last_conclusion).data(1).value = value1
If is_relation(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
     C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
      0, 0, 0, 0, 0, 0, value1, tn%, -2000, 0, 0, 0, _
       con_relation(last_conclusion).data(0), 0, 0, conclusion_data(last_conclusion).ty, _
        record_0.data0.condition_data, 1) Then

Call con_relation_(last_conclusion, con_relation(last_conclusion).data(0))
  conclusion_data(last_conclusion).no(0) = tn%
Else
If con_relation(last_conclusion).data(0).value = "1" Then
 If conclusion_data(last_conclusion).ty = eline_ Then
 con_eline(last_conclusion).data(0).data0.poi(0) = con_relation(last_conclusion).data(0).poi(0)
 con_eline(last_conclusion).data(0).data0.poi(1) = con_relation(last_conclusion).data(0).poi(1)
 con_eline(last_conclusion).data(0).data0.poi(2) = con_relation(last_conclusion).data(0).poi(2)
 con_eline(last_conclusion).data(0).data0.poi(3) = con_relation(last_conclusion).data(0).poi(3)
 con_eline(last_conclusion).data(0).data0.n(0) = con_relation(last_conclusion).data(0).n(0)
 con_eline(last_conclusion).data(0).data0.n(1) = con_relation(last_conclusion).data(0).n(1)
 con_eline(last_conclusion).data(0).data0.n(2) = con_relation(last_conclusion).data(0).n(2)
 con_eline(last_conclusion).data(0).data0.n(3) = con_relation(last_conclusion).data(0).poi(3)
 con_eline(last_conclusion).data(0).data0.line_no(0) = con_relation(last_conclusion).data(0).line_no(0)
 con_eline(last_conclusion).data(0).data0.line_no(0) = con_relation(last_conclusion).data(0).line_no(0)
 ElseIf conclusion_data(last_conclusion).ty = midpoint_ Then
 con_mid_point(last_conclusion).data(0).poi(0) = con_relation(last_conclusion).data(0).poi(0)
 con_mid_point(last_conclusion).data(0).poi(1) = con_relation(last_conclusion).data(0).poi(1)
 con_mid_point(last_conclusion).data(0).poi(2) = con_relation(last_conclusion).data(0).poi(3)
 con_mid_point(last_conclusion).data(0).n(0) = con_relation(last_conclusion).data(0).n(0)
 con_mid_point(last_conclusion).data(0).n(1) = con_relation(last_conclusion).data(0).n(1)
 con_mid_point(last_conclusion).data(0).n(2) = con_relation(last_conclusion).data(0).n(3)
 con_mid_point(last_conclusion).data(0).line_no = con_relation(last_conclusion).data(0).line_no(0)
 End If
End If
End If
i% = line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                conclusion, conclusion_color, 1, 0)
j% = line_number(C_display_wenti.m_point_no(num, 2), _
                 C_display_wenti.m_point_no(num, 3), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                 conclusion, conclusion_color, 1, 0)
'End If
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
Call set_Drelation(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
    C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
      0, 0, 0, 0, 0, 0, value1, temp_record, 0, 0, 0, 0, 0, True)
End If
End Sub

Public Function draw_picture_17_18(ByVal num As Integer, ByVal no_reduce As Byte) As Boolean
''-17 △□□□是等腰直角三角形
'-18 △□□□是等腰三角形
Dim i%, u%, v%
Dim tl(2) As Integer
Dim A&
Dim r1!, r2!
For i% = 0 To 1
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
       C_display_wenti.m_condition(num, i%)) Then
  draw_picture_17_18 = True
Exit Function
End If
Next i%
Call line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                 condition, conclusion_color, 1, 0)
If C_display_wenti.m_no(num) = -18 Then
If draw_free_point(C_display_wenti.m_point_no(num, 2), _
       C_display_wenti.m_condition(num, 2)) Then
  draw_picture_17_18 = True
Exit Function
End If
   r1! = sqr((m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X - _
            m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X) ^ 2 + _
          (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y - _
            m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y) ^ 2)
   r2! = sqr((m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X - _
            m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.X) ^ 2 + _
          (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y - _
            m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.Y) ^ 2)
        Call set_point_visible(C_display_wenti.m_point_no(num, 2), 0, True)
   'Call draw_point(Draw_form, poi(C_display_wenti.m_point_no(num,2)), 0, delete)
    t_coord.X = m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X + _
          (m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.X - _
            m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X) * r1! / r2!
    t_coord.Y = m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y + _
          (m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.Y - _
            m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y) * r1! / r2!
            Call set_point_coordinate(C_display_wenti.m_point_no(num, 2), _
                   t_coord, False)
   'Call draw_point(Draw_form, poi(C_display_wenti.m_point_no(num,2)), 0, display)
  '       tl(0) = line_number(c_display_wenti.m_point_no(num,1), _
                   c_display_wenti.m_point_no(num,0), condition, display)
         tl(1) = line_number(C_display_wenti.m_point_no(num, 0), _
                             C_display_wenti.m_point_no(num, 2), _
                             pointapi0, pointapi0, _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                             condition, condition_color, 1, 0)
         tl(2) = line_number(C_display_wenti.m_point_no(num, 1), _
                             C_display_wenti.m_point_no(num, 2), _
                             pointapi0, pointapi0, _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                             condition, condition_color, 1, 0)

Else
If draw_free_point(C_display_wenti.m_point_no(num, 2), _
       C_display_wenti.m_condition(num, 2)) Then
  draw_picture_17_18 = True
Exit Function
End If
    u% = m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X - _
         m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X
    v% = m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y - _
         m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y
    A& = m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X * _
          m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y + _
           m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X * _
            m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.Y + _
             m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.X * _
              m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y - _
               m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y * _
                m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X - _
                 m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y * _
                  m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.X - _
                   m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.Y * _
                    m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X
          Call set_point_visible(C_display_wenti.m_point_no(num, 2), 0, True)
    ' Call draw_point(Draw_form, m_poi(C_display_wenti.m_point_no(num,2)), 0, delete)
     If A& > 0 Then
       Call C_display_wenti.set_m_point_no(num, 0, 10, False)
       t_coord.X = m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X - v%
       t_coord.Y = m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y + u%
       Call set_point_coordinate(C_display_wenti.m_point_no(num, 2), _
              t_coord, False)
 Else
    Call C_display_wenti.set_m_point_no(num, 1, 10, False)
       t_coord.X = m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X + v%
       t_coord.Y = m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y - u%
        Call set_point_coordinate(C_display_wenti.m_point_no(num, 2), t_coord, False)
 End If
  m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree = 0
    tl(0) = line_number(C_display_wenti.m_point_no(num, 1), _
                        C_display_wenti.m_point_no(num, 0), _
                        pointapi0, pointapi0, _
                        depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                        depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                        condition, condition_color, 1, 0)
    tl(1) = line_number(C_display_wenti.m_point_no(num, 0), _
                        C_display_wenti.m_point_no(num, 2), _
                        pointapi0, pointapi0, _
                        depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                        depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                        condition, condition_color, 1, 0)
    tl(2) = line_number(C_display_wenti.m_point_no(num, 1), _
                        C_display_wenti.m_point_no(num, 2), _
                        pointapi0, pointapi0, _
                        depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                        depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                        condition, condition_color, 1, 0)
    'Call  draw_point(Draw_form, poi(C_display_wenti.m_point_no(num,2)), 0, display)
 End If
If C_display_wenti.m_no(num) = -17 Then
 Call vertical_line(tl(0), tl(1), True, True)
Else
 Call paral_line(tl(0), tl(1), True, True)
End If
End Function

Public Sub set_name_for_draw_picture0()
Dim i%
Dim num As Integer
For num = 1 To C_display_wenti.m_last_input_wenti_no
For i% = 0 To 10

  If Asc(C_display_wenti.m_condition(num, i%)) > 63 And _
           Asc(C_display_wenti.m_condition(num, i%)) < 91 Then
     If set_or_prove < 2 Then
    ' If is_used_char(num, i%) = False Then
     '  last_used_char = last_used_char + 1
      '  used_char(last_used_char) = c_display_wenti.m_condition(num,i%)

     'End If
    Else
    Call C_display_wenti.set_m_point_no(num, _
     point_number(C_display_wenti.m_condition(num, i%)), i%, True)
         '读出条件的点, 记录点号
    End If

  End If
Next i%
Next num
End Sub

Public Function draw_picture18_21(ByVal num As Integer, ByVal no_reduce As Byte) As Boolean
Dim i%, m_i1%, m_i2%
Dim triA(2) As Integer
Dim tp(2) As Integer
Dim temp_record As total_record_type
For i% = 1 To 3
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
      C_display_wenti.m_condition(num, i%)) Then
       draw_picture18_21 = True
        Exit Function
End If
  If i% <> 3 Then
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   End If '点poi(c_display_wenti.m_point_no(num,num,i%))参加推理
Next i%
temp_record.record_.display_no = -(num + 1)
If C_display_wenti.m_point_no(num, 0) = 0 Then
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
 Call C_display_wenti.set_m_point_no(num, _
      last_conditions.last_cond(1).point_no, 0, False)
End If
MDIForm1.Toolbar1.Buttons(21).Image = 33
'poi(last_conditions.last_cond(1).point_no).data(0).data0.name = C_display_wenti.m_condition(num,0)
m_poi(last_conditions.last_cond(1).point_no).data(0).degree = 0
Call set_point_visible(last_conditions.last_cond(1).point_no, 1, False)
'End If
If C_display_wenti.m_no(num) = 18 Then
 Call centroid(C_display_wenti.m_point_no(num, 1), _
C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), _
    C_display_wenti.m_point_no(num, 0), _
       C_display_wenti.m_point_no(num, 4), _
          C_display_wenti.m_point_no(num, 5), _
             C_display_wenti.m_point_no(num, 6), True)
ElseIf C_display_wenti.m_no(num) = 19 Then
  temp_circle(0) = circumcenter(C_display_wenti.m_point_no(num, 1), _
    C_display_wenti.m_point_no(num, 2), _
     C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 0)) ',
'      poi(c_display_wenti.m_point_no(num,num,0)).data(0).data0.name = c_display_wenti.m_condition(num,0)
'       Call put_name(c_display_wenti.m_point_no(num,num,0))
     'For i% = 1 To 3
     ' Call add_point_to_circle(C_display_wenti.m_point_no(num,num,i%), temp_circle(0), True)
     'Next i%
    Call C_display_wenti.set_m_point_no(num, temp_circle(0), 10, False)
ElseIf C_display_wenti.m_no(num) = 20 Then
Call orthocenter(C_display_wenti.m_point_no(num, 1), _
  C_display_wenti.m_point_no(num, 2), _
   C_display_wenti.m_point_no(num, 3), _
     C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 4), _
       C_display_wenti.m_point_no(num, 5), _
         C_display_wenti.m_point_no(num, 6), 0, True)
   temp_line(0) = line_number0(C_display_wenti.m_point_no(num, 0), _
       C_display_wenti.m_point_no(num, 1), 0, 0)
   temp_line(1) = line_number0(C_display_wenti.m_point_no(num, 2), _
       C_display_wenti.m_point_no(num, 3), 0, 0)
   Call set_dverti(temp_line(0), temp_line(1), temp_record, 0, 0, True)
   temp_line(0) = line_number0(C_display_wenti.m_point_no(num, 0), _
       C_display_wenti.m_point_no(num, 2), 0, 0)
   temp_line(1) = line_number0(C_display_wenti.m_point_no(num, 3), _
       C_display_wenti.m_point_no(num, 1), 0, 0)
   Call set_dverti(temp_line(0), temp_line(1), temp_record, 0, 0, True)
   temp_line(0) = line_number0(C_display_wenti.m_point_no(num, 0), _
       C_display_wenti.m_point_no(num, 3), 0, 0)
   temp_line(1) = line_number0(C_display_wenti.m_point_no(num, 1), _
       C_display_wenti.m_point_no(num, 2), 0, 0)
   Call set_dverti(temp_line(0), temp_line(1), temp_record, 0, 0, True)
'       Call put_name(c_display_wenti.m_point_no(num,num,0))
ElseIf C_display_wenti.m_no(num) = 21 Then
Call incenter(C_display_wenti.m_point_no(num, 1), _
C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 10), _
  C_display_wenti.m_point_no(num, 4), C_display_wenti.m_point_no(num, 5), _
     C_display_wenti.m_point_no(num, 6), False) '4圆567切点
End If
'poi(c_display_wenti.m_point_no(num,num,0)).data(0).data0.name = c_display_wenti.m_condition(num,0)
 'Call put_name(c_display_wenti.m_point_no(num,num,0))
    For i% = 4 To 6
If C_display_wenti.m_point_no(num, i%) > 0 Then
 Call get_new_char(C_display_wenti.m_point_no(num, i%))
  '  Call put_name(C_display_wenti.m_point_no(num,num,i%))
 'End If
 End If
   Next i%
'End If
End Function
Public Sub draw_picture_2(ByVal num As Integer, ByVal no_reduce As Byte)
'-2 作⊙□[down\\(_)]和⊙□[down\\(_)]的公切线□□
Dim i%, j%, k%
For i% = 0 To 3
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
      C_display_wenti.m_condition(num, i%)) = True Then
        Exit Sub
End If
Next i%
i% = m_circle_number(1, C_display_wenti.m_point_no(num, 0), pointapi0, _
        C_display_wenti.m_point_no(num, 1), 0, 0, 0, 0, 0, _
         1, 1, condition, condition_color, True)
j% = m_circle_number(1, C_display_wenti.m_point_no(num, 2), pointapi0, _
        C_display_wenti.m_point_no(num, 3), 0, 0, 0, 0, 0, _
         1, 1, condition, condition_color, True)
Call tangent_line_for_two_circle(C_display_wenti.m_point_no(num, 7), _
   C_display_wenti.m_point_no(num, 8), no_reduce)
For k% = 2 To 5
Call draw_t_line(k%)
Next k%
draw_picture1_mark_2:
   event_statue = wait_for_draw_point
    While event_statue = wait_for_draw_point
      DoEvents
    Wend
 If event_statue = draw_point_down Or _
      event_statue = draw_point_move Or _
         event_statue = draw_point_up Then 'mouse_type <> 1 Then
    t_coord = input_coord
    'temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark_2
 End If

 Call C_display_wenti.set_m_point_no(num, _
      read_T_line1(t_coord), 10, True)
If C_display_wenti.m_point_no(num, 10) > 1 Then
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
MDIForm1.Toolbar1.Buttons(21).Image = 33
'   Call init_Point0(last_conditions.last_cond(1).point_no)
temp_point(0).no = last_conditions.last_cond(1).point_no
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
MDIForm1.Toolbar1.Buttons(21).Image = 33
 '  Call init_Point0(last_conditions.last_cond(1).point_no)
temp_point(1).no = last_conditions.last_cond(1).point_no
'poi(temp_point(0)).data(0).data0.coordinate.X = tangent_line(C_display_wenti.m_point_no(num,num,10)).p(0).X
'poi(temp_point(0)).data(0).data0.coordinate.Y = tangent_line(C_display_wenti.m_point_no(num,num,10)).p(0).Y
Call set_point_coordinate(temp_point(0).no, tangent_line(C_display_wenti.m_point_no(num, 10)).p(0), False)
'poi(temp_point(1)).data(0).data0.coordinate.X = tangent_line(C_display_wenti.m_point_no(num,num,10)).p(1).X
'poi(temp_point(1)).data(0).data0.coordinate.Y = tangent_line(C_display_wenti.m_point_no(num,num,10)).p(1).Y
Call set_point_coordinate(temp_point(1).no, tangent_line(C_display_wenti.m_point_no(num, 10)).p(1), False)
temp_line(1) = line_number(temp_point(0).no, temp_point(1).no, _
                           pointapi0, pointapi0, _
                           depend_condition(point_, temp_point(0).no), _
                           depend_condition(point_, temp_point(1).no), _
                           condition, condition_color, 1, 0)
temp_line(2) = line_number(temp_point(0).no, m_Circ(temp_circle(0)).data(0).data0.center, _
                           pointapi0, pointapi0, _
                           depend_condition(point_, temp_point(0).no), _
                           depend_condition(point_, m_Circ(temp_circle(0)).data(0).data0.center), _
                           condition, condition_color, 0, 0)
temp_line(3) = line_number(temp_point(0).no, m_Circ(temp_circle(1)).data(0).data0.center, _
                           pointapi0, pointapi0, _
                           depend_condition(point_, temp_point(0).no), _
                           depend_condition(point_, m_Circ(temp_circle(1)).data(0).data0.center), _
                           condition, condition_color, 0, 0)
'Call add_point_to_circle(temp_point(0), C_display_wenti.m_point_no(num,num,7), True)
'Call add_point_to_circle(temp_point(1), C_display_wenti.m_point_no(num,num,8), True)
'poi(temp_point(0)).data(0).data0.name = C_display_wenti.m_condition(num,4)
'poi(temp_point(1)).data(0).data0.name = C_display_wenti.m_condition(num,5)
Call C_display_wenti.set_m_point_no(num, temp_point(0).no, 4, True)
Call C_display_wenti.set_m_point_no(num, temp_point(1).no, 5, True)
Call set_point_visible(temp_point(0).no, 1, False)
Call set_point_visible(temp_point(1).no, 1, False)
'Call draw_point(Draw_form, poi(temp_point(0)), 0, display)
'Call draw_point(Draw_form, poi(temp_point(1)), 0, display)
Call vertical_line(temp_line(1), temp_line(2), True, True)
Call vertical_line(temp_line(1), temp_line(3), True, True)
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
MDIForm1.Toolbar1.Buttons(21).Image = 33
'   Call init_Point0(last_conditions.last_cond(1).point_no)
 Call get_new_char(last_conditions.last_cond(1).point_no)
  Call C_display_wenti.set_m_point_no(num, _
       last_conditions.last_cond(1).point_no, 15, False)
   t_coord.X = _
      2 * m_poi(C_display_wenti.m_point_no(num, 4)).data(0).data0.coordinate.X - _
          m_poi(C_display_wenti.m_point_no(num, 5)).data(0).data0.coordinate.X
   t_coord.Y = _
      2 * m_poi(C_display_wenti.m_point_no(num, 4)).data(0).data0.coordinate.Y - _
          m_poi(C_display_wenti.m_point_no(num, 5)).data(0).data0.coordinate.Y
   Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
   record_0.data0.condition_data.condition_no = 0
   Call add_point_to_line(last_conditions.last_cond(1).point_no, temp_line(1), 0, no_display, _
        True, 0, temp_record)
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
MDIForm1.Toolbar1.Buttons(21).Image = 33
  ' Call init_Point0(last_conditions.last_cond(1).point_no)
 Call get_new_char(last_conditions.last_cond(1).point_no)
   Call C_display_wenti.set_m_point_no(num, _
         last_conditions.last_cond(1).point_no, 16, False)
   t_coord.X = _
       2 * m_poi(C_display_wenti.m_point_no(num, 5)).data(0).data0.coordinate.X - _
           m_poi(C_display_wenti.m_point_no(num, 4)).data(0).data0.coordinate.X
   t_coord.Y = _
       2 * m_poi(C_display_wenti.m_point_no(num, 5)).data(0).data0.coordinate.Y - _
           m_poi(C_display_wenti.m_point_no(num, 4)).data(0).data0.coordinate.Y
  Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
  record_0.data0.condition_data.condition_no = 0
 Call add_point_to_line(last_conditions.last_cond(1).point_no, temp_line(1), 0, no_display, _
      True, 0, temp_record)
For k% = 2 To 5
Call draw_t_line(k%)
Next k%
Else
    GoTo draw_picture1_mark_2
End If
'end If
'End If
End Sub

Public Sub draw_picture39(ByVal num%)
'39□□、□□、□□三直线共点
Dim temp_record As total_record_type
Dim i%, j%, p%, l1%, l2%, l3%
l1% = line_number(C_display_wenti.m_point_no(num, 0), _
                  C_display_wenti.m_point_no(num, 1), _
                  pointapi0, pointapi0, _
                  depend_condition(0, 0), depend_condition(0, 0), _
                  conclusion, conclusion_color, 1, 0)
l2% = line_number(C_display_wenti.m_point_no(num, 2), _
                  C_display_wenti.m_point_no(num, 3), _
                  pointapi0, pointapi0, _
                  depend_condition(0, 0), depend_condition(0, 0), _
                  conclusion, conclusion_color, 1, 0)
l3% = line_number(C_display_wenti.m_point_no(num, 4), _
                  C_display_wenti.m_point_no(num, 5), _
                  pointapi0, pointapi0, _
                  depend_condition(0, 0), depend_condition(0, 0), _
                  conclusion, conclusion_color, 1, 0)
For i% = 1 To last_conditions.last_cond(1).point_no
 If is_point_in_line3(i%, m_lin(l1%).data(0).data0, 0) And _
    is_point_in_line3(i%, m_lin(l2%).data(0).data0, 0) Then
   p% = i%
    GoTo draw_picture39_mark0
 ElseIf is_point_in_line3(i%, m_lin(l2%).data(0).data0, 0) And _
    is_point_in_line3(i%, m_lin(l3%).data(0).data0, 0) Then
   p% = i%
    Call exchange_two_integer(l1%, l3%)
    GoTo draw_picture39_mark0
 ElseIf is_point_in_line3(i%, m_lin(l1%).data(0).data0, 0) And _
    is_point_in_line3(i%, m_lin(l3%).data(0).data0, 0) Then
   p% = i%
     Call exchange_two_integer(l2%, l3%)
   GoTo draw_picture39_mark0
 End If
Next i%
GoTo draw_picture39_mark2
draw_picture39_mark0:
For i% = 0 To num
For j% = 0 To 6
If C_display_wenti.m_point_no(i%, j%) = p% Then
GoTo draw_picture39_mark1
End If
Next j%
Next i%
GoTo draw_picture39_mark3
draw_picture39_mark2:
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
MDIForm1.Toolbar1.Buttons(21).Image = 33
  ' Call init_Point0(last_conditions.last_cond(1).point_no)
p% = last_conditions.last_cond(1).point_no
draw_picture39_mark3:
Call inter_point_line_line3(C_display_wenti.m_point_no(num, 0), paral_, _
        l1%, m_lin(l2%).data(0).data0.poi(0), True, l2%, t_coord, p%, False, _
          temp_record.record_data.data0.condition_data, True)
'Call get_new_char
Call set_point_name(p%, next_char(p%, "", 0, 0))
'Call put_name(p%)
record_0.data0.condition_data.condition_no = 0
Call add_point_to_line(p%, l1%, 0, no_display, True, 0, temp_record)
Call add_point_to_line(p%, l2%, 0, no_display, True, 0, temp - record)
If last_conditions.last_cond(1).new_point_no Mod 10 = 0 Then
ReDim Preserve new_point(last_conditions.last_cond(1).new_point_no + 10) As new_point_type
End If
 last_conditions.last_cond(1).new_point_no = last_conditions.last_cond(1).new_point_no + 1
 temp_record.record_data.data0.condition_data.condition_no = 1 ' record0
 temp_record.record_data.data0.condition_data.condition(1).ty = new_point_  ' record0
 temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(1).new_point_no ' record0
' temp_record.record_data.data1.aid_condition = last_conditions.last_cond(1).new_point_no
new_point(last_conditions.last_cond(1).new_point_no).data(0) = new_point_data_0
new_point(last_conditions.last_cond(1).new_point_no).data(0).poi(0) = p%
   new_point(last_conditions.last_cond(1).new_point_no).data(0).display_string = LoadResString_(1730, _
      "\\1\" + m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.name + _
                m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.name + _
      "\\2\\" + m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.name + _
                m_poi(C_display_wenti.m_point_no(num, 3)).data(0).data0.name + _
      "\\3\\" + m_poi(p%).data(0).data0.name)
temp_record.record_data.data0.condition_data.condition_no = last_conditions.last_cond(1).new_point_no
 temp_record.record_data.data0.condition_data.condition(1).ty = new_point_
  temp_record.record_data.data0.condition_data.condition(1).no = last_conditions.last_cond(0).new_point_no
   'temp_record.record_data.record_dataaid_condition = 0
draw_picture39_mark1:
  conclusion_data(last_conclusion).ty = point3_on_line_
   con_Three_point_on_line(last_conclusion).data(0).poi(0) = m_lin(l3%).data(0).data0.poi(0) 'c_display_wenti.m_point_no(num,4)
    con_Three_point_on_line(last_conclusion).data(0).poi(1) = p%
     con_Three_point_on_line(last_conclusion).data(0).poi(2) = m_lin(l3%).data(0).data0.poi(1) 'c_display_wenti.m_point_no(num,5)
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True

End Sub

Public Sub draw_picture40(num As Integer)
'△□□□是等边三角形
Dim tpol As polygon
Dim temp_record As total_record_type
tpol.total_v = 3
 tpol.v(0) = C_display_wenti.m_point_no(num, 0)
 tpol.v(1) = C_display_wenti.m_point_no(num, 1)
 tpol.v(2) = C_display_wenti.m_point_no(num, 2)
If set_or_prove < 2 Then
Call draw_triangle(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   C_display_wenti.m_point_no(num, 2), conclusion)
conclusion_data(last_conclusion).ty = epolygon_
Call is_epolygon(tpol, 0, con_Epolygon(last_conclusion).data(0))
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
Call set_Epolygon(tpol, temp_record, 0, 0, 0, 0, False) '-1)
End If
End Sub

Public Sub draw_picture25_27_28(ByVal num As Integer) 'num 语句号
'25□□＝□□
'27□□∥□□
'28□□⊥□□
Dim i%, j%, k%, l%, n%
Dim n_(7) As Integer
Dim tp(3) As Integer
Dim tl(1) As Integer
Dim ty1 As Byte
Dim dn(2) As Integer
Dim t_n(1) As Integer
Dim con_ty As Byte
Dim el_data As eline_data0_type
Dim temp_record As total_record_type
Dim c_data As condition_data_type
For i% = 0 To 3
  If i% <> 3 Then
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
  End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
'***************************************************************************
'画线段
tl(0) = line_number(C_display_wenti.m_point_no(num, 0), _
                    C_display_wenti.m_point_no(num, 1), _
                    pointapi0, pointapi0, _
                    depend_condition(0, 0), depend_condition(0, 0), _
                    conclusion, conclusion_color, 1, 0)
tl(1) = line_number(C_display_wenti.m_point_no(num, 2), _
                    C_display_wenti.m_point_no(num, 3), _
                    pointapi0, pointapi0, _
                    depend_condition(0, 0), depend_condition(0, 0), _
                    conclusion, conclusion_color, 1, 0)
'***********************************************************************************
'***************************************************************************
'25□□＝□□
If C_display_wenti.m_no(num) = 25 Then
'For i% = 0 To 3
'    m_poi(C_display_wenti.m_point_no(num,i%)).data(0).no_reduce = 0
'Next i%
conclusion_data(last_conclusion).ty = eline_ '结论类型□□＝□□
record_0.data0.condition_data.condition_no = 0 ' record0
If is_equal_dline(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
      C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
        0, 0, 0, 0, 0, 0, dn(0), n_(0), n_(1), n_(2), n_(3), el_data, _
           dn(1), dn(2), con_ty, "", record_0.data0.condition_data) Then
           '如果
 If (con_ty = eline_ And dn(0) > 0) Then
 conclusion_data(last_conclusion).no(0) = dn(0)
 con_eline(last_conclusion).data(1).data0.poi(0) = Deline(dn(0)).data(0).data0.poi(0)
 con_eline(last_conclusion).data(1).data0.poi(1) = Deline(dn(0)).data(0).data0.poi(1)
 con_eline(last_conclusion).data(1).data0.poi(2) = Deline(dn(0)).data(0).data0.poi(2)
 con_eline(last_conclusion).data(1).data0.poi(3) = Deline(dn(0)).data(0).data0.poi(3)
 con_eline(last_conclusion).data(1).data0.n(0) = Deline(dn(0)).data(0).data0.n(0)
 con_eline(last_conclusion).data(1).data0.n(1) = Deline(dn(0)).data(0).data0.n(1)
 con_eline(last_conclusion).data(1).data0.n(2) = Deline(dn(0)).data(0).data0.n(2)
 con_eline(last_conclusion).data(1).data0.n(3) = Deline(dn(0)).data(0).data0.n(3)
 con_eline(last_conclusion).data(1).data0.line_no(0) = Deline(dn(0)).data(0).data0.line_no(0)
 con_eline(last_conclusion).data(1).data0.line_no(1) = Deline(dn(0)).data(0).data0.line_no(1)
 ElseIf con_ty = line_value_ Or _
   is_same_two_point(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
     C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3)) Then
 If last_conditions.last_cond(1).eline_no Mod 10 = 0 Then
  ReDim Preserve Deline(last_conditions.last_cond(1).eline_no + 10) As eline_type
 End If
  last_conditions.last_cond(1).eline_no = last_conditions.last_cond(1).eline_no + 1
  Deline(last_conditions.last_cond(1).eline_no).data(0) = eline_data_0
   Deline(last_conditions.last_cond(1).eline_no).data(0).data0 = el_data
    If con_ty = line_value_ Then
    Call add_conditions_to_record(line_value_, dn(0), dn(1), dn(2), _
               temp_record.record_data.data0.condition_data)
    End If
    temp_record.record_.no_reduce = 2
   Deline(last_conditions.last_cond(1).eline_no).data(0).record = temp_record.record_data
   Deline(last_conditions.last_cond(1).eline_no).record_ = temp_record.record_
   For i% = 0 To 3
    For j% = last_conditions.last_cond(1).eline_no To n_(i%) + 2 Step -1
     Deline(j%).data(0).record.data1.index.i(i%) = Deline(j% - 1).data(0).record.data1.index.i(i%)
    Next j%
     Deline(n_(i%) + 1).data(0).record.data1.index.i(i%) = last_conditions.last_cond(1).eline_no
   Next i%
       conclusion_data(last_conclusion).no(0) = last_conditions.last_cond(1).eline_no
 ElseIf con_ty = midpoint_ Then
   conclusion_data(last_conclusion).ty = midpoint_
     conclusion_data(last_conclusion).no(0) = dn(0)
     c_data.condition_no = 0
    Call is_mid_point(el_data.poi(0), el_data.poi(1), el_data.poi(3), _
      el_data.n(0), el_data.n(1), el_data.n(3), el_data.line_no(0), _
       0, -2000, 0, 0, 0, 0, 0, 0, con_mid_point(last_conclusion).data(0), _
        "", 0, 0, 0, c_data)

 End If
Else
If con_ty = midpoint_ Then
conclusion_data(last_conclusion).ty = midpoint_
c_data.condition_no = 0
Call is_mid_point(el_data.poi(0), el_data.poi(1), el_data.poi(3), _
      el_data.n(0), el_data.n(1), el_data.n(3), el_data.line_no(0), _
       0, -2000, 0, 0, 0, 0, 0, 0, con_mid_point(last_conclusion).data(0), _
        "", 0, 0, 0, c_data)
Else
conclusion_data(last_conclusion).ty = eline_
con_eline(last_conclusion).data(1).data0.poi(0) = C_display_wenti.m_point_no(num, 0)
con_eline(last_conclusion).data(1).data0.poi(1) = C_display_wenti.m_point_no(num, 1)
con_eline(last_conclusion).data(1).data0.poi(2) = C_display_wenti.m_point_no(num, 2)
con_eline(last_conclusion).data(1).data0.poi(3) = C_display_wenti.m_point_no(num, 3)
con_eline(last_conclusion).data(0).data0 = el_data
End If
End If
ElseIf C_display_wenti.m_no(num) = 27 Then
'tl(0) = line_number0(C_display_wenti.m_point_no(num, 0), _
                    C_display_wenti.m_point_no(num, 1), _
                     0, 0)
'tl(1) = line_number0(C_display_wenti.m_point_no(num, 2), _
                    C_display_wenti.m_point_no(num, 3), _
                     0, 0)
conclusion_data(last_conclusion).ty = paral_
If is_dparal(tl(0), tl(1), n%, n_(0), n_(1), n_(2), _
    con_paral(last_conclusion).data(0).line_no(0), _
        con_paral(last_conclusion).data(0).line_no(1)) Then
If n% > 0 Then
conclusion_data(last_conclusion).no(0) = n%
ElseIf tl(0) = tl(1) Then
If last_conditions.last_cond(1).paral_no Mod 10 = 0 Then
ReDim Preserve Dparal(last_conditions.last_cond(1).paral_no + 10) As paral_type
End If
last_conditions.last_cond(1).paral_no = last_conditions.last_cond(1).paral_no + 1
Dparal(last_conditions.last_cond(1).paral_no).data(0).data0 = two_line_data_0
Dparal(last_conditions.last_cond(1).paral_no).data(0).data0.line_no(0) = tl(0)
 Dparal(last_conditions.last_cond(1).paral_no).data(0).data0.line_no(1) = tl(1)
  Dparal(last_conditions.last_cond(1).paral_no).record_.no_reduce = 7
   Dparal(last_conditions.last_cond(1).paral_no).data(0).data0.record = temp_record.record_data
    Dparal(last_conditions.last_cond(1).paral_no).record_ = temp_record.record_
    For i% = 0 To 2
     For j% = last_conditions.last_cond(1).paral_no To n_(i%) - 2 Step -1
      Dparal(j%).data(0).data0.record.data1.index.i(i%) = Dparal(j% - 1).data(0).data0.record.data1.index.i(i%)
     Next j%
      Dparal(n_(i%) + 1).data(0).data0.record.data1.index.i(i%) = last_conditions.last_cond(1).paral_no
    Next i%
   conclusion_data(last_conclusion).no(0) = last_conditions.last_cond(1).paral_no
End If
End If
Else
'tl(0) = line_number0(C_display_wenti.m_point_no(num, 0), _
                    C_display_wenti.m_point_no(num, 1), _
                     0, 0)
'tl(1) = line_number0(C_display_wenti.m_point_no(num, 2), _
                    C_display_wenti.m_point_no(num, 3), _
                     0, 0)
conclusion_data(last_conclusion).ty = verti_
If is_dverti(tl(0), tl(1), n%, n_(0), n_(1), n_(2), _
               con_verti(last_conclusion).data(0).line_no(0), _
                 con_verti(last_conclusion).data(0).line_no(1)) Then
conclusion_data(last_conclusion).no(0) = n%
End If
End If
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
'Else
End Sub

Public Sub draw_picture45(ByVal num As Integer)
'45 □□□□是平行四边形
Dim i%
Dim j%
Dim tp(3) As Integer
tp(0) = C_display_wenti.m_point_no(num, 0)
tp(1) = C_display_wenti.m_point_no(num, 1)
tp(2) = C_display_wenti.m_point_no(num, 2)
tp(3) = C_display_wenti.m_point_no(num, 3)
j% = 0
For i% = 1 To 3
If tp(j%) > tp(i%) Then
j% = i%
End If
Next i%
If tp(1) > tp(3) Then
Call exchange_two_integer(tp(1), tp(3))
End If
If set_or_prove < 2 Then
Call draw_polygon4(tp(0), tp(1), tp(2), tp(3), conclusion)
End If
conclusion_data(last_conclusion).ty = parallelogram_
For i% = 0 To 3
con_parallelogram(last_conclusion).data(0).poi(i%) = tp(i%)
Next i%
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
''Else
'End If
End Sub
Public Sub draw_picture49(ByVal num As Integer)
'□□□□是等腰梯形
Dim temp_record As total_record_type
If set_or_prove < 2 Then
Call draw_polygon4(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   C_display_wenti.m_point_no(num, 2), _
                   C_display_wenti.m_point_no(num, 3), conclusion)
conclusion_data(last_conclusion).ty = tixing_
con_Dtixing(last_conclusion).data(0).poly4_no = polygon4_number( _
  C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
   C_display_wenti.m_point_no(num, 3), 0)
 con_Dtixing(last_conclusion).data(0).ty = equal_side_tixing_
If Dpolygon4(con_Dtixing(last_conclusion).data(0).poly4_no%).data(0).poi(0) = _
    C_display_wenti.m_point_no(num, 0) Or _
   Dpolygon4(con_Dtixing(last_conclusion).data(0).poly4_no%).data(0).poi(0) = _
    C_display_wenti.m_point_no(num, 2) Then
     con_Dtixing(last_conclusion).data(0).poi(0) = _
          Dpolygon4(con_Dtixing(last_conclusion).data(0).poly4_no%).data(0).poi(0)
     con_Dtixing(last_conclusion).data(0).poi(1) = _
          Dpolygon4(con_Dtixing(last_conclusion).data(0).poly4_no%).data(0).poi(1)
     con_Dtixing(last_conclusion).data(0).poi(2) = _
          Dpolygon4(con_Dtixing(last_conclusion).data(0).poly4_no%).data(0).poi(2)
     con_Dtixing(last_conclusion).data(0).poi(3) = _
          Dpolygon4(con_Dtixing(last_conclusion).data(0).poly4_no%).data(0).poi(3)
ElseIf Dpolygon4(con_Dtixing(last_conclusion).data(0).poly4_no%).data(0).poi(0) = _
        C_display_wenti.m_point_no(num, 1) Or _
       Dpolygon4(con_Dtixing(last_conclusion).data(0).poly4_no%).data(0).poi(0) = _
        C_display_wenti.m_point_no(num, 3) Then
     con_Dtixing(last_conclusion).data(0).poi(0) = _
          Dpolygon4(con_Dtixing(last_conclusion).data(0).poly4_no%).data(0).poi(1)
     con_Dtixing(last_conclusion).data(0).poi(1) = _
          Dpolygon4(con_Dtixing(last_conclusion).data(0).poly4_no%).data(0).poi(2)
     con_Dtixing(last_conclusion).data(0).poi(2) = _
          Dpolygon4(con_Dtixing(last_conclusion).data(0).poly4_no%).data(0).poi(3)
     con_Dtixing(last_conclusion).data(0).poi(3) = _
          Dpolygon4(con_Dtixing(last_conclusion).data(0).poly4_no%).data(0).poi(0)
End If
conclusion_data(last_conclusion).wenti_no = num
  last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
record_0.data0.condition_data.condition_no = 0 ' record0
Call set_equal_dline(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 1), _
  C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, 0, 0, _
   temp_record, 0, 0, 0, 0, 0, True)
record_0.data0.condition_data.condition_no = 0 ' record0
Call set_dparal(line_number0(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), 0, 0), line_number0(C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), 0, 0), temp_record, 0, 3, False)
End If

End Sub

Public Sub draw_picture44(ByVal num As Integer)
'44□□□□是正方形
Dim tpol As polygon
Dim n%
Dim temp_record As total_record_type
Dim poly4_no%
conclusion_data(last_conclusion).ty = epolygon_
tpol.total_v = 4
 tpol.v(0) = C_display_wenti.m_point_no(num, 0)
 tpol.v(1) = C_display_wenti.m_point_no(num, 1)
 tpol.v(2) = C_display_wenti.m_point_no(num, 2)
 tpol.v(3) = C_display_wenti.m_point_no(num, 3)
If set_or_prove < 2 Then
Call draw_polygon4(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   C_display_wenti.m_point_no(num, 2), _
                   C_display_wenti.m_point_no(num, 3), conclusion)
     poly4_no% = polygon4_number(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
       C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), 0)
If is_epolygon(tpol, n%, con_Epolygon(last_conclusion).data(0)) Then
conclusion_data(last_conclusion).ty = epolygon_
conclusion_data(last_conclusion).no(0) = n%
Else
     poly4_no% = polygon4_number(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
       C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), 0)
        If is_squre0(poly4_no%, n%, 0) Then
         conclusion_data(last_conclusion).ty = Squre
         conclusion_data(last_conclusion).no(0) = n%
        Else
        con_squre(last_conclusion).data(0).polygon4_no = poly4_no%
               conclusion_data(last_conclusion).ty = Squre
       End If
End If
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
record_0.data0.condition_data.condition_no = 0 'record0
Call set_squre(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
                C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
                 0, temp_record, 1, False) '-1)
End If
End Sub





Public Sub draw_picture51(ByVal num As Integer)
'51 ∠□□□=∠□□□+∠□□□
Dim A(2) As Integer
Dim temp_record As total_record_type
conclusion_data(last_conclusion).ty = angle3_value_
A(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 0), _
                        C_display_wenti.m_point_no(num, 1), _
                        C_display_wenti.m_point_no(num, 2), 0, 0))
A(1) = Abs(angle_number(C_display_wenti.m_point_no(num, 3), _
                        C_display_wenti.m_point_no(num, 4), _
                        C_display_wenti.m_point_no(num, 5), 0, 0))
A(2) = Abs(angle_number(C_display_wenti.m_point_no(num, 6), _
                        C_display_wenti.m_point_no(num, 7), _
                        C_display_wenti.m_point_no(num, 8), 0, 0))
If set_or_prove < 2 Then
Call draw_angle(C_display_wenti.m_point_no(num, 0), _
                C_display_wenti.m_point_no(num, 1), _
                C_display_wenti.m_point_no(num, 2), conclusion)
Call draw_angle(C_display_wenti.m_point_no(num, 3), _
                C_display_wenti.m_point_no(num, 4), _
                C_display_wenti.m_point_no(num, 5), conclusion)
Call draw_angle(C_display_wenti.m_point_no(num, 6), _
                C_display_wenti.m_point_no(num, 7), _
                C_display_wenti.m_point_no(num, 8), conclusion)
temp_record.record_data.data0.condition_data.condition_no = 0
con_angle3_value(last_conclusion).data(1).data0.angle(0) = A(0)
con_angle3_value(last_conclusion).data(1).data0.angle(1) = A(1)
con_angle3_value(last_conclusion).data(1).data0.angle(2) = A(2)
con_angle3_value(last_conclusion).data(1).data0.para(0) = "1"
con_angle3_value(last_conclusion).data(1).data0.para(1) = "-1"
con_angle3_value(last_conclusion).data(1).data0.para(2) = "-1"
con_angle3_value(last_conclusion).data(1).data0.value = "0"
con_angle3_value(last_conclusion).data(1).data0.value = "0"
Call is_three_angle_value(A(0), A(1), A(2), _
  "1", "-1", "-1", "0", "0", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
     con_angle3_value(last_conclusion).data(0).data0, temp_record.record_data.data0.condition_data, 0)
conclusion_data(last_conclusion).wenti_no = num
  last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
 MDIForm1.Toolbar1.Buttons(19).visible = True
Else
record_0.data0.condition_data.condition_no = 0 ' record0
Call set_three_angle_value(A(0), A(1), A(2), "1", "-1", "-1", _
 "0", 0, temp_record, 0, 0, 0, 0, 0, 0, False)
End If

End Sub

Public Sub draw_picture52(ByVal num As Integer)
'52 ∠□□□+∠□□□=!_~ °
Dim A(1) As Integer
Dim i%, bra%
Dim temp_record As total_record_type
Dim value1 As String
A(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 0), _
                        C_display_wenti.m_point_no(num, 1), _
                        C_display_wenti.m_point_no(num, 2), 0, 0))
A(1) = Abs(angle_number(C_display_wenti.m_point_no(num, 3), _
                        C_display_wenti.m_point_no(num, 4), _
                        C_display_wenti.m_point_no(num, 5), 0, 0))
i% = 6
While Asc(C_display_wenti.m_condition(num, i%)) > 13 ' c_display_wenti.m_condition(i%) <> empty_char
 If C_display_wenti.m_condition(num, i%) < "A" Then
value1 = value1 + C_display_wenti.m_condition(num, i%)
i% = i% + 1
 Else
 value1 = value1 + C_display_wenti.m_condition(num, i%)
 i% = i% + 1
End If
Wend
If set_or_prove < 2 Then
conclusion_data(last_conclusion).ty = angle3_value_
Call draw_angle(C_display_wenti.m_point_no(num, 0), _
                C_display_wenti.m_point_no(num, 1), _
                C_display_wenti.m_point_no(num, 2), conclusion)
Call draw_angle(C_display_wenti.m_point_no(num, 3), _
                C_display_wenti.m_point_no(num, 4), _
                C_display_wenti.m_point_no(num, 5), conclusion)
    con_angle3_value(last_conclusion).data(1).data0.angle(0) = A(0)
     con_angle3_value(last_conclusion).data(1).data0.angle(1) = A(1)
      con_angle3_value(last_conclusion).data(1).data0.angle(2) = 0
    con_angle3_value(last_conclusion).data(1).data0.para(0) = "1"
     con_angle3_value(last_conclusion).data(1).data0.para(1) = "1"
      con_angle3_value(last_conclusion).data(1).data0.para(2) = "0"
        con_angle3_value(last_conclusion).data(1).data0.value = value1
        Call is_three_angle_value(A(0), A(1), 0, "1", "1", "0", value1, value1, 0, 0, 0, _
          -2000, 0, 0, 0, 0, 0, 0, 0, con_angle3_value(last_conclusion).data(0).data0, _
            temp_record.record_data.data0.condition_data, 0)
conclusion_data(last_conclusion).wenti_no = num
              last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
 MDIForm1.Toolbar1.Buttons(19).visible = True
Else
Call set_three_angle_value(A(0), A(1), 0, "1", "1", "0", value1, _
        0, temp_record, 0, 0, 0, 0, 0, 0, False)
End If
End Sub
Public Sub draw_picture46(num As Integer)
'46□□□□是菱形
Dim temp_record As total_record_type
Dim p4 As Integer
If set_or_prove < 2 Then
Call draw_polygon4(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   C_display_wenti.m_point_no(num, 2), _
                   C_display_wenti.m_point_no(num, 3), conclusion)
conclusion_data(last_conclusion).ty = rhombus_
Call is_rhombus(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
      C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), 0, p4%, 0, 0)
con_rhombus(last_conclusion).data(0).polygon4_no = p4%
conclusion_data(last_conclusion).wenti_no = num
 last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
'record_0 = record0
Call set_rhombus(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), temp_record, 0, 3)
End If
End Sub

Public Sub draw_picture43(ByVal num As Integer)
'43□□□□是长方形
Dim temp_record As total_record_type
If set_or_prove < 2 Then
Call draw_polygon4(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   C_display_wenti.m_point_no(num, 2), _
                   C_display_wenti.m_point_no(num, 3), conclusion)
conclusion_data(last_conclusion).ty = long_squre_
 con_long_squre(last_conclusion).data(0).polygon4_no = polygon4_number( _
               C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
                C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), 0)
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
'record_0 = record0
Call set_long_squre(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), temp_record, 0, 3, 1, False)
End If
End Sub

Public Sub draw_picture41(ByVal num As Integer)
'41 △□□□是等腰三角形
Dim temp_record As total_record_type
Dim triA%
If set_or_prove < 2 Then
conclusion_data(last_conclusion).ty = equal_side_triangle_
triA% = triangle_number(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
 C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, 0, con_equal_side_triangle(last_conclusion).data(0).direction)
Call draw_triangle(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   C_display_wenti.m_point_no(num, 2), conclusion)
con_equal_side_triangle(last_conclusion).data(0).triangle = triA%
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
record_0.data0.condition_data.condition_no = 0 ' record0
Call set_equal_dline(C_display_wenti.m_point_no(num, 0), _
  C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 0), _
   C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, 0, 0, _
    temp_record, 0, 0, 0, 0, 0, True)
End If
End Sub

Public Sub draw_picture42(ByVal num As Integer)
'42△□□□是等腰直角三角形
Dim temp_record As total_record_type
Dim triA%
If set_or_prove < 2 Then
conclusion_data(last_conclusion).ty = equal_side_right_triangle_
triA% = triangle_number(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
 C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, 0, con_equal_side_right_triangle(last_conclusion).data(0).direction)
Call draw_triangle(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   C_display_wenti.m_point_no(num, 2), conclusion)
con_equal_side_right_triangle(last_conclusion).data(0).triangle = triA%
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
record_0.data0.condition_data.condition_no = 0 ' record0
Call set_equal_dline(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 0), _
  C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, 0, 0, _
   temp_record, 0, 0, 0, 0, 0, True)
'record_0 = record0
Call set_angle_value(Abs(angle_number(C_display_wenti.m_point_no(num, 1), _
 C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 2), 0, 0)), _
   "90", temp_record, 0, 0, False)
End If
End Sub
Public Sub draw_temp_equal_sides_triangle(ByVal p1%, ByVal p2%, m_coord As POINTAPI, _
                           coord1 As POINTAPI, out_p1%)
Dim r1!
Dim r2!
   r1! = sqr((m_poi(p1%).data(0).data0.coordinate.X - _
            m_poi(p2%).data(0).data0.coordinate.X) ^ 2 + _
          (m_poi(p1%).data(0).data0.coordinate.Y - _
            m_poi(p2%).data(0).data0.coordinate.Y) ^ 2)
   r2! = sqr((m_poi(p1%).data(0).data0.coordinate.X - _
            m_coord.X) ^ 2 + _
          (m_poi(p1%).data(0).data0.coordinate.Y - _
            m_coord.Y) ^ 2)
    coord1.X = _
         m_poi(p1%).data(0).data0.coordinate.X + _
          (m_coord.X - _
            m_poi(p1%).data(0).data0.coordinate.X) * r1! / r2!
    coord1.Y = _
         m_poi(p1%).data(0).data0.coordinate.Y + _
          (m_coord.Y - _
            m_poi(p1%).data(0).data0.coordinate.Y) * r1! / r2!
If out_p1% > 0 Then
   Call set_point_coordinate(out_p1%, t_coord1, False)
End If
End Sub

Public Sub draw_temp_long_squre(m_coord As POINTAPI, ty As Integer)
Call set_temp_long_squre(m_coord, ty)
Draw_form.Line (temp_four_point_fig.p(1).X, temp_four_point_fig.p(1).Y)- _
      (temp_four_point_fig.p(2).X, temp_four_point_fig.p(2).Y), QBColor(fill_color)
Draw_form.Line (temp_four_point_fig.p(2).X, temp_four_point_fig.p(2).Y)- _
      (temp_four_point_fig.p(3).X, temp_four_point_fig.p(3).Y), QBColor(fill_color)
Draw_form.Line (temp_four_point_fig.p(3).X, temp_four_point_fig.p(3).Y)- _
      (temp_four_point_fig.p(0).X, temp_four_point_fig.p(0).Y), QBColor(fill_color)
End Sub

Public Sub set_temp_long_squre(m_coord As POINTAPI, ty As Integer)
Dim t!
t! = CSng(((m_coord.X - temp_four_point_fig.p(0).X) * (temp_four_point_fig.p(1).X - _
             temp_four_point_fig.p(0).X) + _
             (m_coord.Y - temp_four_point_fig.p(0).Y) * (temp_four_point_fig.p(1).Y - _
             temp_four_point_fig.p(0).Y))) / _
             ((temp_four_point_fig.p(1).X - temp_four_point_fig.p(0).X) ^ 2 + _
              (temp_four_point_fig.p(1).Y - temp_four_point_fig.p(0).Y) ^ 2)
'm_coord 在p1%,p2% 上的投影
temp_four_point_fig.p(3).X = m_coord.X - t! * _
       (temp_four_point_fig.p(1).X - temp_four_point_fig.p(0).X)
temp_four_point_fig.p(3).Y = m_coord.Y - t! * _
     (temp_four_point_fig.p(1).Y - temp_four_point_fig.p(0).Y)
temp_four_point_fig.p(2).X = m_coord.X + (1 - t!) * _
     (temp_four_point_fig.p(1).X - temp_four_point_fig.p(0).X)
temp_four_point_fig.p(2).Y = m_coord.Y + (1 - t!) * _
     (temp_four_point_fig.p(1).Y - temp_four_point_fig.p(0).Y)
If temp_four_point_fig.poi(2) > 0 Then
   Call set_point_coordinate(temp_four_point_fig.poi(2), temp_four_point_fig.p(2), True)
End If
If temp_four_point_fig.poi(3) > 0 Then
   Call set_point_coordinate(temp_four_point_fig.poi(3), temp_four_point_fig.p(3), True)
End If
End Sub

Public Sub draw_temp_parallelogram(m_coord As POINTAPI, ty As Integer)
Call set_temp_parallelogram(m_coord, ty)
Draw_form.Line (temp_four_point_fig.p(1).X, temp_four_point_fig.p(1).Y)- _
      (temp_four_point_fig.p(2).X, temp_four_point_fig.p(2).Y), QBColor(fill_color)
Draw_form.Line (temp_four_point_fig.p(0).X, temp_four_point_fig.p(0).Y)- _
      (temp_four_point_fig.p(3).X, temp_four_point_fig.p(3).Y), QBColor(fill_color)
Draw_form.Line (temp_four_point_fig.p(2).X, temp_four_point_fig.p(2).Y)- _
      (temp_four_point_fig.p(3).X, temp_four_point_fig.p(3).Y), QBColor(fill_color)
End Sub

Public Sub set_temp_parallelogram(m_coord As POINTAPI, ty As Integer)
If ty = 0 Or ty = 1 Then
temp_four_point_fig.p(2) = m_coord
 temp_four_point_fig.p(3).X = temp_four_point_fig.p(2).X + _
       (temp_four_point_fig.p(0).X - temp_four_point_fig.p(1).X)
 temp_four_point_fig.p(3).Y = temp_four_point_fig.p(2).Y + _
       (temp_four_point_fig.p(0).Y - temp_four_point_fig.p(1).Y)
Else
temp_four_point_fig.p(3) = m_coord
 temp_four_point_fig.p(2).X = temp_four_point_fig.p(3).X + _
       (temp_four_point_fig.p(1).X - temp_four_point_fig.p(0).X)
 temp_four_point_fig.p(2).Y = temp_four_point_fig.p(3).Y + _
       (temp_four_point_fig.p(1).Y - temp_four_point_fig.p(0).Y)
End If
If temp_four_point_fig.poi(2) > 0 Then
   Call set_point_coordinate(temp_four_point_fig.poi(2), temp_four_point_fig.p(2), True)
End If
If temp_four_point_fig.poi(3) > 0 Then
   Call set_point_coordinate(temp_four_point_fig.poi(3), temp_four_point_fig.p(3), True)
End If
End Sub

Public Sub set_temp_equal_side_tixing(m_coord As POINTAPI, ty As Integer)
Dim t!
Dim r&
temp_four_point_fig.p(2) = m_coord
r& = (temp_four_point_fig.p(0).X - temp_four_point_fig.p(1).X) ^ 2 + _
       (temp_four_point_fig.p(0).Y - temp_four_point_fig.p(1).Y) ^ 2
If ty = 0 Then
If r& > 0 Then
t! = CSng((temp_four_point_fig.p(0).X - temp_four_point_fig.p(1).X) * _
        (temp_four_point_fig.p(1).X + temp_four_point_fig.p(0).X - 2 * temp_four_point_fig.p(2).X) + _
       (temp_four_point_fig.p(0).Y - temp_four_point_fig.p(1).Y) * _
        (temp_four_point_fig.p(1).Y + temp_four_point_fig.p(0).Y - 2 * temp_four_point_fig.p(2).Y)) / r&
temp_four_point_fig.p(3).X = temp_four_point_fig.p(2).X + _
        t! * (temp_four_point_fig.p(0).X - temp_four_point_fig.p(1).X)
temp_four_point_fig.p(3).Y = temp_four_point_fig.p(2).Y + _
        t! * (temp_four_point_fig.p(0).Y - temp_four_point_fig.p(1).Y)
Else
temp_four_point_fig.p(3) = temp_four_point_fig.p(2)
' Y2& = Y1&
End If
Else
If r& > 0 Then
t! = CSng((temp_four_point_fig.p(1).X - temp_four_point_fig.p(0).X) * _
        (temp_four_point_fig.p(0).X + temp_four_point_fig.p(1).X - 2 * temp_four_point_fig.p(3).X) + _
       (temp_four_point_fig.p(1).Y - temp_four_point_fig.p(0).Y) * _
        (temp_four_point_fig.p(0).Y + temp_four_point_fig.p(1).Y - 2 * temp_four_point_fig.p(3).Y)) / r&
temp_four_point_fig.p(2).X = temp_four_point_fig.p(3).X + _
        t! * (temp_four_point_fig.p(1).X - temp_four_point_fig.p(0).X)
temp_four_point_fig.p(2).Y = temp_four_point_fig.p(3).Y + _
        t! * (temp_four_point_fig.p(1).Y - temp_four_point_fig.p(0).Y)
Else
temp_four_point_fig.p(2) = temp_four_point_fig.p(3)
' Y2& = Y1&
End If
End If
If temp_four_point_fig.poi(2) > 0 Then
   Call set_point_coordinate(temp_four_point_fig.poi(2), temp_four_point_fig.p(2), True)
End If
If temp_four_point_fig.poi(3) > 0 Then
   Call set_point_coordinate(temp_four_point_fig.poi(3), temp_four_point_fig.p(3), True)
End If
End Sub
Public Sub draw_temp_equal_side_tixing(m_coord As POINTAPI, ty As Integer)
Call set_temp_equal_side_tixing(m_coord, ty)
Draw_form.Line (temp_four_point_fig.p(1).X, temp_four_point_fig.p(1).Y)- _
      (temp_four_point_fig.p(2).X, temp_four_point_fig.p(2).Y), QBColor(fill_color)
Draw_form.Line (temp_four_point_fig.p(0).X, temp_four_point_fig.p(0).Y)- _
      (temp_four_point_fig.p(3).X, temp_four_point_fig.p(3).Y), QBColor(fill_color)
Draw_form.Line (temp_four_point_fig.p(2).X, temp_four_point_fig.p(2).Y)- _
      (temp_four_point_fig.p(3).X, temp_four_point_fig.p(3).Y), QBColor(fill_color)
End Sub

Public Sub draw_temp_rhombus(m_coord As POINTAPI, ty As Integer)
Call set_temp_rhombus(m_coord, ty)
Draw_form.Line (temp_four_point_fig.p(1).X, temp_four_point_fig.p(1).Y)- _
      (temp_four_point_fig.p(2).X, temp_four_point_fig.p(2).Y), QBColor(fill_color)
Draw_form.Line (temp_four_point_fig.p(0).X, temp_four_point_fig.p(0).Y)- _
      (temp_four_point_fig.p(3).X, temp_four_point_fig.p(3).Y), QBColor(fill_color)
Draw_form.Line (temp_four_point_fig.p(2).X, temp_four_point_fig.p(2).Y)- _
      (temp_four_point_fig.p(3).X, temp_four_point_fig.p(3).Y), QBColor(fill_color)
End Sub

Public Sub set_temp_rhombus(m_coord As POINTAPI, ty As Integer)
Dim r1&, r2&
r1& = sqr((temp_four_point_fig.p(1).X - temp_four_point_fig.p(0).X) ^ 2 + _
       (temp_four_point_fig.p(1).Y - temp_four_point_fig.p(0).Y) ^ 2)
If ty = 0 Then
r2& = sqr((temp_four_point_fig.p(1).X - m_coord.X) ^ 2 + _
       (temp_four_point_fig.p(1).Y - m_coord.Y) ^ 2)
If r2& = 0 Then
temp_four_point_fig.p(2).X = 2 * temp_four_point_fig.p(1).X - temp_four_point_fig.p(0).X
temp_four_point_fig.p(2).Y = 2 * temp_four_point_fig.p(1).Y - temp_four_point_fig.p(0).Y
temp_four_point_fig.p(3) = temp_four_point_fig.p(1)
Else
temp_four_point_fig.p(2).X = temp_four_point_fig.p(1).X + (m_coord.X - temp_four_point_fig.p(1).X) * _
      r1& / r2&
temp_four_point_fig.p(2).Y = temp_four_point_fig.p(1).Y + (m_coord.Y - temp_four_point_fig.p(1).Y) * _
      r1& / r2&
temp_four_point_fig.p(3).X = temp_four_point_fig.p(0).X + _
         (temp_four_point_fig.p(2).X - temp_four_point_fig.p(1).X)
temp_four_point_fig.p(3).Y = temp_four_point_fig.p(0).Y + _
         (temp_four_point_fig.p(2).Y - temp_four_point_fig.p(1).Y)
End If
Else
r2& = sqr((temp_four_point_fig.p(0).X - m_coord.X) ^ 2 + _
       (temp_four_point_fig.p(0).Y - m_coord.Y) ^ 2)
If r2& = 0 Then
temp_four_point_fig.p(3).X = 2 * temp_four_point_fig.p(0).X - temp_four_point_fig.p(1).X
temp_four_point_fig.p(3).Y = 2 * temp_four_point_fig.p(0).Y - temp_four_point_fig.p(1).Y
temp_four_point_fig.p(2) = temp_four_point_fig.p(0)
Else
temp_four_point_fig.p(3).X = temp_four_point_fig.p(0).X + (m_coord.X - temp_four_point_fig.p(0).X) * _
      r1& / r2&
temp_four_point_fig.p(3).Y = temp_four_point_fig.p(0).Y + (m_coord.Y - temp_four_point_fig.p(0).Y) * _
      r1& / r2&
temp_four_point_fig.p(2).X = temp_four_point_fig.p(1).X + _
         (temp_four_point_fig.p(3).X - temp_four_point_fig.p(0).X)
temp_four_point_fig.p(2).Y = temp_four_point_fig.p(1).Y + _
         (temp_four_point_fig.p(3).Y - temp_four_point_fig.p(0).Y)
End If
End If
If temp_four_point_fig.poi(2) > 0 Then
   Call set_point_coordinate(temp_four_point_fig.poi(2), temp_four_point_fig.p(2), True)
End If
If temp_four_point_fig.poi(3) > 0 Then
   Call set_point_coordinate(temp_four_point_fig.poi(3), temp_four_point_fig.p(3), True)
End If
End Sub

Public Function draw_picture_15(ByVal num As Integer, ByVal no_reduce As Byte) As Boolean
'-15 □□□□是梯形
Dim i%
Dim ele1 As condition_type
Dim ele2 As condition_type
For i% = 0 To 2
 If draw_free_point(C_display_wenti.m_point_no(num, i%), _
       C_display_wenti.m_condition(num, i%)) = True Then
     draw_picture_15 = True
      Exit Function
 End If
Next i%
Call draw_polygon4(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   C_display_wenti.m_point_no(num, 2), _
                   C_display_wenti.m_point_no(num, 3), condition)
If C_display_wenti.m_point_no(num, 3) = 0 Then
'last_set_point = last_set_point + 1
'   temp_point(6) = 100 - last_set_point
'       Call set_set_point(temp_point(6))
     t_coord.X = m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.X + _
      m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X - m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X
     t_coord.Y = m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.Y + _
      m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y - m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y
       temp_point(6).no = 0
       Call set_aid_point(temp_point(6).no, t_coord, 1)
       temp_line(1) = line_number0(C_display_wenti.m_point_no(num, 2), _
                       temp_point(6).no, 0, 0)
        'lin(temp_line(1)).data(0).data0.visible = 4
draw_picture_15_mark0:
 event_statue = wait_for_draw_point  '输点状态
   While event_statue = wait_for_draw_point '等待事件发生
    DoEvents
   Wend
If event_statue = draw_point_down Or event_statue = _
             draw_point_move Or event_statue = _
                    draw_point_up Then 'mouse_type <> 1 Then
    t_coord = input_coord
    'temp_y& = input_coord.Y
    input_point_type% = read_inter_point(t_coord, ele1, _
                                   ele2, temp_point(0).no, True)
          Call set_point_no_reduce(temp_point(0).no, 0)
     If input_point_type% = new_point_on_line And _
         ele1.no = temp_line(1) Then
      Call C_display_wenti.set_m_point_no(num, temp_point(0).no, 3, True)
      Call set_point_name(C_display_wenti.m_point_no(num, 3), _
            C_display_wenti.m_condition(num, 3))
      Call set_point_visible(C_display_wenti.m_point_no(num, 3), 1, False)
      Call draw_polygon4(C_display_wenti.m_point_no(num, 0), _
                         C_display_wenti.m_point_no(num, 1), _
                         C_display_wenti.m_point_no(num, 2), _
                         C_display_wenti.m_point_no(num, 3), condition)
      'Call put_name(C_display_wenti.m_point_no(num,3))
       temp_point(0).no = 0
        temp_line(1) = 0
      Call remove_point(temp_point(6).no, display, 0)
       Call C_display_wenti.set_m_point_no(num, 1, 4, False)
      If Abs(m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X - _
           m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X) > 5 Then
        Call C_display_wenti.set_m_point_no(num, ( _
          m_poi(C_display_wenti.m_point_no(num, 3)).data(0).data0.coordinate.X - _
            m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.X) * 1000 / _
              (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X - _
                m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X), 5, False)
      Else
        Call C_display_wenti.set_m_point_no(num, ( _
             m_poi(C_display_wenti.m_point_no(num, 3)).data(0).data0.coordinate.Y - _
               m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.Y) * 1000 / _
                 (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y - _
                    m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y), 5, False)
      End If
     Else
       GoTo draw_picture_15_mark0
     End If
 ElseIf event_statue = wait_for_input_char Or _
      event_statue = wait_for_modify_char Then
  draw_picture_15 = True
   Exit Function
 Else
     GoTo draw_picture_15_mark0
 End If
 End If
End Function


Public Sub draw_picture56(ByVal num As Integer)
'求△□□□的面积
Dim A As Integer
Dim i%, n%
Dim value1 As String
Dim triA As triangle_data0_type
Dim temp_record As total_record_type
conclusion_data(last_conclusion).ty = area_of_element_
record_0.data0.condition_data.condition_no = 0 'record0
Call set_triangle(C_display_wenti.m_point_no(num, 0), _
       C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
         triA, A, 0, 0, 0, 0, temp_record, 0)
 con_Area_of_element(last_conclusion).data(0).element.no = A
  con_Area_of_element(last_conclusion).data(0).element.ty = triangle_
  Call draw_triangle(C_display_wenti.m_point_no(num, 0), _
                     C_display_wenti.m_point_no(num, 1), _
                     C_display_wenti.m_point_no(num, 2), conclusion)
If is_area_of_element(triangle_, A, n%, -1000) Then
  conclusion_data(last_conclusion).no(0) = n%
End If
  conclusion_data(last_conclusion).ty = area_of_element_
 conclusion_data(last_conclusion).wenti_no = num
           last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
area_of_triangle_conclusion = 1
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
 MDIForm1.Toolbar1.Buttons(19).visible = True
th_chose(155).chose = 1
th_chose(156).chose = 1
End Sub
Public Sub draw_picture59(ByVal num As Integer)
'59 扇形□□□的面积=?
'c_display_wenti.m_point_no(num,0) 是圆心
Dim i%
conclusion_data(last_conclusion).ty = area_of_fan_
For i% = 1 To C_display_picture.m_circle.Count
 If m_Circ(i%).data(0).data0.center = C_display_wenti.m_point_no(num, 1) Then
    con_Area_of_fan(last_conclusion).data(0).poi(0) = C_display_wenti.m_point_no(num, 0)
     con_Area_of_fan(last_conclusion).data(0).poi(1) = C_display_wenti.m_point_no(num, 1)
      con_Area_of_fan(last_conclusion).data(0).poi(2) = C_display_wenti.m_point_no(num, 2)
       GoTo darw_picture59_mark0
 ElseIf m_Circ(i%).data(0).data0.center = C_display_wenti.m_point_no(num, 0) Then
    con_Area_of_fan(last_conclusion).data(0).poi(0) = C_display_wenti.m_point_no(num, 1)
     con_Area_of_fan(last_conclusion).data(0).poi(1) = C_display_wenti.m_point_no(num, 0)
      con_Area_of_fan(last_conclusion).data(0).poi(2) = C_display_wenti.m_point_no(num, 2)
       GoTo darw_picture59_mark0
 ElseIf m_Circ(i%).data(0).data0.center = C_display_wenti.m_point_no(num, 2) Then
    con_Area_of_fan(last_conclusion).data(0).poi(0) = C_display_wenti.m_point_no(num, 1)
     con_Area_of_fan(last_conclusion).data(0).poi(1) = C_display_wenti.m_point_no(num, 2)
      con_Area_of_fan(last_conclusion).data(0).poi(2) = C_display_wenti.m_point_no(num, 0)
       GoTo darw_picture59_mark0
 End If
Next i%
darw_picture59_mark0:
conclusion_data(last_conclusion).wenti_no = num
        last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
End Sub
Public Sub draw_picture57(num As Integer)
'57 求四边形□□□□的面积
Dim A As Integer
Dim i%, tn%
Dim value1 As String
conclusion_data(last_conclusion).ty = area_of_element_
con_Area_of_element(last_conclusion).data(0).element.no = _
  polygon4_number(C_display_wenti.m_point_no(num, 0), _
           C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
             C_display_wenti.m_point_no(num, 3), 0)
con_Area_of_element(last_conclusion).data(0).element.ty = polygon_
Call draw_polygon4(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   C_display_wenti.m_point_no(num, 2), _
                   C_display_wenti.m_point_no(num, 3), conclusion)
If is_area_of_element(polygon_, _
      con_Area_of_element(last_conclusion).data(0).element.no, _
        tn%, -1000) Then
     conclusion_data(last_conclusion).no(0) = tn%
End If
conclusion_data(last_conclusion).wenti_no = num
      last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
th_chose(155).chose = 1
th_chose(156).chose = 1
End Sub

Public Sub draw_picture60_62(num As Integer)
'△□□□的周长=?
Dim A As Integer
Dim i%
Dim tl(2) As Integer
Dim tn(2, 1) As Integer
Dim value1 As String
Dim value2 As String
Dim triA As triangle_data0_type
Dim temp_record As total_record_type
If C_display_wenti.m_no(num) = 62 Then
i% = 3
While Asc(C_display_wenti.m_condition(num, i%)) > 13
 If C_display_wenti.m_condition(num, i%) < "A" And _
      C_display_wenti.m_condition(num, i%) > "Z" Then
  value1 = value1 + C_display_wenti.m_condition(num, i%)
   i% = i% + 1
 Else
 value2 = value2 + C_display_wenti.m_condition(num, i%)
 i% = i% + 1
End If
Wend
If value2 <> "" And value1 <> "" Then
 value1 = value1 + "*" + value2
ElseIf value1 = "" Then
 value1 = value2
End If
value1 = value_string(value1)
con_length_of_polygon(last_conclusion).data(0).Area = value1
End If
record_0.data0.condition_data.condition_no = 0 'record0
Call set_triangle(C_display_wenti.m_point_no(num, 0), _
       C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
         triA, A, 0, 0, 0, 0, temp_record, 0)
Call draw_triangle(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   C_display_wenti.m_point_no(num, 2), conclusion)
tl(2) = line_number0(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
        tn(2, 0), tn(2, 1))
tl(0) = line_number0(C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
        tn(0, 0), tn(0, 1))
tl(1) = line_number0(C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 0), _
        tn(1, 0), tn(1, 1))
If tn(0, 0) > tn(0, 1) Then
 Call exchange_two_integer(tn(0, 0), tn(0, 1))
End If
If tn(1, 0) > tn(1, 1) Then
 Call exchange_two_integer(tn(1, 0), tn(1, 1))
End If
If tn(2, 0) > tn(2, 1) Then
 Call exchange_two_integer(tn(2, 0), tn(2, 1))
End If
If tl(0) > tl(1) Then
Call exchange_two_integer(tl(0), tl(1))
Call exchange_two_integer(tn(0, 0), tn(1, 0))
Call exchange_two_integer(tn(0, 1), tn(1, 1))
End If
If tl(1) > tl(2) Then
Call exchange_two_integer(tl(1), tl(2))
Call exchange_two_integer(tn(1, 0), tn(2, 0))
Call exchange_two_integer(tn(1, 1), tn(2, 1))
End If
If tl(0) > tl(1) Then
Call exchange_two_integer(tl(0), tl(1))
Call exchange_two_integer(tn(0, 0), tn(1, 0))
Call exchange_two_integer(tn(0, 1), tn(1, 1))
End If
con_length_of_polygon(last_conclusion).polygon_ty = triangle_
con_length_of_polygon(last_conclusion).polygon_no = A
con_length_of_polygon(last_conclusion).data(0).last_segment = 0
For i% = tn(0, 0) To tn(0, 1) - 1
con_length_of_polygon(last_conclusion).data(0).last_segment = _
     con_length_of_polygon(last_conclusion).data(0).last_segment + 1
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).line_no = tl(0)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).n(0) = i%
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).n(1) = i% + 1
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).poi(0) = m_lin(tl(0)).data(0).data0.in_point(i%)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).poi(1) = m_lin(tl(0)).data(0).data0.in_point(i% + 1)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).para = "1"
Next i%
For i% = tn(1, 0) To tn(1, 1) - 1
con_length_of_polygon(last_conclusion).data(0).last_segment = _
     con_length_of_polygon(last_conclusion).data(0).last_segment + 1
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).line_no = tl(1)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).n(0) = i%
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).n(1) = i% + 1
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).poi(0) = _
       m_lin(tl(1)).data(0).data0.in_point(i%)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).poi(1) = _
       m_lin(tl(1)).data(0).data0.in_point(i% + 1)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).para = "1"
Next i%
For i% = tn(2, 0) To tn(2, 1) - 1
con_length_of_polygon(last_conclusion).data(0).last_segment = _
     con_length_of_polygon(last_conclusion).data(0).last_segment + 1
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).line_no = tl(2)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).n(0) = i%
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).n(1) = i% + 1
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).poi(0) = _
       m_lin(tl(2)).data(0).data0.in_point(i%)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).poi(1) = _
       m_lin(tl(2)).data(0).data0.in_point(i% + 1)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).para = "1"
Next i%
con_length_of_polygon(last_conclusion).record_.conclusion_no = last_conclusion + 1
    conclusion_data(last_conclusion).ty = length_of_polygon_
 temp_record.record_.conclusion_no = last_conclusion + 1
Call set_length_of_polygon(con_length_of_polygon(last_conclusion), 0, _
       temp_record)
 conclusion_data(last_conclusion).wenti_no = num
       last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True


End Sub

Public Sub draw_picture61(num As Integer)
'61 求⊙□[down\\(_)]的周长
Dim A As Integer
Dim i%
Dim value1 As String
conclusion_data(last_conclusion).ty = sides_length_of_circle_
A = m_circle_number(1, C_display_wenti.m_point_no(num, 0), pointapi0, _
                    C_display_wenti.m_point_no(num, 1), 0, 0, 0, 0, 0, 1, 1, _
                     conclusion, conclusion_color, True)
    con_Sides_length_of_circle(last_conclusion).data(0).circ = A
 conclusion_data(last_conclusion).wenti_no = num
       last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
End Sub

Public Sub draw_picture55(num As Integer)
'∠□□□/∠□□□=!_~
Dim A1, A2 As Integer
Dim i%
Dim value1 As String
Dim ty As Byte
Dim temp_record As total_record_type
value1 = initial_string(number_string(C_display_wenti.m_point_no(num, 6))) 'initial_string(cond_to_string(num, 6, 18, 0))
'Call read_number_from_wenti(num, 6, 0, 0, value1)
If set_or_prove < 2 Then
A1 = angle_number(C_display_wenti.m_point_no(num, 0), _
                  C_display_wenti.m_point_no(num, 1), _
                  C_display_wenti.m_point_no(num, 2), 0, 0)
A2 = angle_number(C_display_wenti.m_point_no(num, 3), _
                  C_display_wenti.m_point_no(num, 4), _
                  C_display_wenti.m_point_no(num, 5), 0, 0)
Call draw_angle(C_display_wenti.m_point_no(num, 0), _
                C_display_wenti.m_point_no(num, 1), _
                C_display_wenti.m_point_no(num, 2), conclusion)
Call draw_angle(C_display_wenti.m_point_no(num, 3), _
                C_display_wenti.m_point_no(num, 4), _
                C_display_wenti.m_point_no(num, 5), conclusion)
conclusion_data(last_conclusion).ty = angle3_value_
'Call combine_two_angle(Abs(A1), Abs(A2), _
  con_angle3_value(last_conclusion).data(0).angle(0), _
   0, 0, con_angle3_value(last_conclusion).data(0).angle(1), 0, _
     0, ty, 0, 1)
  con_angle3_value(last_conclusion).data(0).data0.angle(0) = Abs(A1)
  con_angle3_value(last_conclusion).data(0).data0.angle(1) = Abs(A2)
   con_angle3_value(last_conclusion).data(0).data0.angle(2) = 0
If ty < 9 Then
  Call ratio_value1(value1, ty, con_angle3_value(last_conclusion).data(0).data0.para(1))
  con_angle3_value(last_conclusion).data(0).data0.para(1) = time_string("-1", _
     con_angle3_value(last_conclusion).data(0).data0.para(1), True, False)
  con_angle3_value(last_conclusion).data(0).data0.para(0) = "1"
  con_angle3_value(last_conclusion).data(0).data0.para(2) = "0"
  con_angle3_value(last_conclusion).data(0).data0.value = "0"
  con_angle3_value(last_conclusion).data(1) = con_angle3_value(last_conclusion).data(0)
  Call is_three_angle_value(con_angle3_value(last_conclusion).data(1).data0.angle(0), _
    con_angle3_value(last_conclusion).data(1).data0.angle(1), con_angle3_value(last_conclusion).data(1).data0.angle(2), _
     con_angle3_value(last_conclusion).data(1).data0.para(0), con_angle3_value(last_conclusion).data(1).data0.para(1), _
      con_angle3_value(last_conclusion).data(1).data0.para(2), con_angle3_value(last_conclusion).data(1).data0.value, _
       con_angle3_value(last_conclusion).data(1).data0.value, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
        con_angle3_value(last_conclusion).data(0).data0, record_0.data0.condition_data, 0)
        If is_uselly_para(con_angle3_value(last_conclusion).data(0).data0.para(0)) = False Or _
            is_uselly_para(con_angle3_value(last_conclusion).data(0).data0.para(1)) = False Or _
              is_uselly_para(con_angle3_value(last_conclusion).data(0).data0.para(2)) = False Then
          is_uselly_para_for_angle = False
        End If
conclusion_data(last_conclusion).wenti_no = num
        last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
record_0.data0.condition_data.condition_no = 0 'record0
Call set_angle_relation(A1, A2, value1, "1", temp_record, 0, 0, True)
End If
End If
End Sub

Public Sub draw_picture47(num As Integer)
'47 直线□□与⊙□[down\\(_)]相切于□
Dim l%, c%
Dim temp_record As total_record_type
l% = line_number0(C_display_wenti.m_point_no(num, 0), _
    C_display_wenti.m_point_no(num, 1), 0, 0)
    c% = m_circle_number(1, C_display_wenti.m_point_no(num, 2), pointapi0, _
      C_display_wenti.m_point_no(num, 3), 0, 0, 0, 0, 0, 1, 1, _
       conclusion, conclusion_color, True)
If event_statue = input_prove_by_hand Then
record_0.data0.condition_data.condition_no = 0 ' record0
Call set_tangent_line(l%, C_display_wenti.m_point_no(num, 4), _
   c%, C_display_wenti.m_point_no(num, 2), 0, temp_record, 0, 3)
Else
con_tangent_line(last_conclusion).data(0).line_no = l%
Call line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), depend_condition(0, 0), _
                 conclusion, conclusion_color, 1, 0)
con_tangent_line(last_conclusion).data(0).poi(0) = C_display_wenti.m_point_no(num, 4)
con_tangent_line(last_conclusion).data(0).poi(1) = C_display_wenti.m_point_no(num, 2)
con_tangent_line(last_conclusion).data(0).circ(0) = c%
con_tangent_line(last_conclusion).data(0).circ(1) = 0
conclusion_data(last_conclusion).ty = tangent_line_
 conclusion_data(last_conclusion).wenti_no = num
         last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
End If
End Sub

Public Sub draw_picture29(num As Integer)
'29点 □位于线段□□的垂直平分线上
Dim l%
Dim temp_record As total_record_type
If set_or_prove < 2 Then
Call is_equal_dline(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
     C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, 0, 0, _
      -2000, 0, 0, 0, con_eline(last_conclusion).data(1).data0, 0, 0, 0, "", _
        temp_record.record_data.data0.condition_data)
 con_eline(last_conclusion).data(0) = con_eline(last_conclusion).data(1)
'l% = line_number(c_display_wenti.m_point_no(num,1), c_display_wenti.m_point_no(num,2), _
   concl, display)
'Call is_verti_mid_line(c_display_wenti.m_point_no(num,1), _
          0, c_display_wenti.m_point_no(num,2), 0, 0, 0, 0, _
            con_verti_mid_line(last_conclusion).data(0).data0)
conclusion_data(last_conclusion).ty = eline_
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
record_0.data0.condition_data.condition_no = 0 'record0
Call set_verti_mid_line(C_display_wenti.m_point_no(num, 1), _
              C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 2), 0, temp_record, -1, 0)
End If
End Sub
Public Sub draw_picture_47_48(num As Integer, ByVal no_reduce As Byte)
Call C_display_wenti.Get_wenti(num)
Call draw_any_triangle(wenti_cond0.data)
End Sub

Public Sub draw_picture_45_46(num As Integer, ByVal no_reduce As Byte)
Call C_display_wenti.Get_wenti(num)
Call draw_any_polygon4(wenti_cond0.data)
End Sub


Public Sub draw_picture65(num As Integer)
'65 四边形□□□□的面积=!_~
Dim i%, bra%
Dim value1 As String
Dim value2 As String
Dim triA%
Dim temp_record As total_record_type
i% = 4
While Asc(C_display_wenti.m_condition(num, i%)) > 13
 If C_display_wenti.m_condition(num, i%) < "A" Then
value1 = value1 + C_display_wenti.m_condition(num, i%)
i% = i% + 1
 Else
 value2 = value2 + C_display_wenti.m_condition(num, i%)
 i% = i% + 1
End If
Wend
If value2 <> "" And value1 <> "" Then
 value1 = value1 + "*" + value2
ElseIf value1 = "" Then
 value1 = value2
End If
If event_statue = input_prove_by_hand Then
record_0.data0.condition_data.condition_no = 0 'record0
Call set_area_of_polygon(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), value1, temp_record, 0, 3)
  area_of_triangle_conclusion = 1
Else
con_Area_of_element(last_conclusion).data(0).element.ty = polygon_
con_Area_of_element(last_conclusion).data(0).element.no = _
 polygon4_number(C_display_wenti.m_point_no(num, 0), _
   C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
    C_display_wenti.m_point_no(num, 3), 0)
          area_of_triangle_conclusion = 1
  con_Area_of_element(last_conclusion).data(0).value = value1
   conclusion_data(last_conclusion).ty = area_of_element_
   Call draw_polygon4(C_display_wenti.m_point_no(num, 0), _
                      C_display_wenti.m_point_no(num, 1), _
                      C_display_wenti.m_point_no(num, 2), _
                      C_display_wenti.m_point_no(num, 3), conclusion)
     If is_area_of_element(polygon_, con_Area_of_element(last_conclusion).data(0).element.no, _
                    conclusion_data(last_conclusion).no(0), -1000) = False Then
                     conclusion_data(last_conclusion).no(0) = 0
     End If
 conclusion_data(last_conclusion).wenti_no = num
                    last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
End If

End Sub

Public Sub draw_picture64_66(num As Integer)
'64 四边形□□□□的周长=!_~
'66 四边形□□□□的周长=?
Dim i%, j%, poly4_no%
Dim tl(3) As Integer
Dim tn(3, 1) As Integer
Dim v(3) As String
Dim value1 As String
'Dim value2 As String
Dim it(3) As Integer
Dim n(3) As Integer
'Dim it1(3) As Integer
Dim temp_record  As total_record_type
If C_display_wenti.m_no(num) = 64 Then
 value1 = initial_string(number_string(C_display_wenti.m_point_no(num, 4))) 'initial_string(cond_to_string(num, 4, 18, 0))
  con_length_of_polygon(last_conclusion).data(0).value = value1
End If
 con_length_of_polygon(last_conclusion).data(0).value = "0"
record_0.data0.condition_data.condition_no = 0 'record0
 poly4_no% = polygon4_number(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
                                   C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), 0)
Call draw_polygon4(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   C_display_wenti.m_point_no(num, 2), _
                   C_display_wenti.m_point_no(num, 3), conclusion)
tl(0) = line_number0(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
        tn(0, 0), tn(0, 1))
tl(1) = line_number0(C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
        tn(1, 0), tn(1, 1))
tl(2) = line_number0(C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
        tn(2, 0), tn(2, 1))
tl(3) = line_number0(C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 0), _
        tn(3, 0), tn(3, 1))
If tn(0, 0) > tn(0, 1) Then
 Call exchange_two_integer(tn(0, 0), tn(0, 1))
End If
If tn(1, 0) > tn(1, 1) Then
 Call exchange_two_integer(tn(1, 0), tn(1, 1))
End If
If tn(2, 0) > tn(2, 1) Then
 Call exchange_two_integer(tn(2, 0), tn(2, 1))
End If
If tn(3, 0) > tn(3, 1) Then
 Call exchange_two_integer(tn(3, 0), tn(3, 1))
End If
If tl(0) > tl(1) Then
Call exchange_two_integer(tl(0), tl(1))
Call exchange_two_integer(tn(0, 0), tn(1, 0))
Call exchange_two_integer(tn(0, 1), tn(1, 1))
End If
If tl(1) > tl(2) Then
Call exchange_two_integer(tl(1), tl(2))
Call exchange_two_integer(tn(1, 0), tn(2, 0))
Call exchange_two_integer(tn(1, 1), tn(2, 1))
End If
If tl(2) > tl(3) Then
Call exchange_two_integer(tl(2), tl(3))
Call exchange_two_integer(tn(2, 0), tn(3, 0))
Call exchange_two_integer(tn(2, 1), tn(3, 1))
End If
If tl(0) > tl(1) Then
Call exchange_two_integer(tl(0), tl(1))
Call exchange_two_integer(tn(0, 0), tn(1, 0))
Call exchange_two_integer(tn(0, 1), tn(1, 1))
End If
If tl(1) > tl(2) Then
Call exchange_two_integer(tl(1), tl(2))
Call exchange_two_integer(tn(1, 0), tn(2, 0))
Call exchange_two_integer(tn(1, 1), tn(2, 1))
End If
If tl(0) > tl(1) Then
Call exchange_two_integer(tl(0), tl(1))
Call exchange_two_integer(tn(0, 0), tn(1, 0))
Call exchange_two_integer(tn(0, 1), tn(1, 1))
End If
con_length_of_polygon(last_conclusion).data(0).last_segment = 0
con_length_of_polygon(last_conclusion).polygon_ty = polygon_
con_length_of_polygon(last_conclusion).polygon_no = poly4_no%
For i% = tn(0, 0) To tn(0, 1) - 1
con_length_of_polygon(last_conclusion).data(0).last_segment = _
     con_length_of_polygon(last_conclusion).data(0).last_segment + 1
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).line_no = tl(0)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).n(0) = i%
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).n(1) = i% + 1
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).poi(0) = m_lin(tl(0)).data(0).data0.in_point(i%)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).poi(1) = m_lin(tl(0)).data(0).data0.in_point(i% + 1)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).para = "1"
Next i%
For i% = tn(1, 0) To tn(1, 1) - 1
con_length_of_polygon(last_conclusion).data(0).last_segment = _
     con_length_of_polygon(last_conclusion).data(0).last_segment + 1
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).line_no = tl(1)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).n(0) = i%
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).n(1) = i% + 1
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).poi(0) = _
       m_lin(tl(1)).data(0).data0.in_point(i%)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).poi(1) = _
       m_lin(tl(1)).data(0).data0.in_point(i% + 1)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).para = "1"
Next i%
For i% = tn(2, 0) To tn(2, 1) - 1
con_length_of_polygon(last_conclusion).data(0).last_segment = _
     con_length_of_polygon(last_conclusion).data(0).last_segment + 1
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).line_no = tl(2)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).n(0) = i%
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).n(1) = i% + 1
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).poi(0) = _
       m_lin(tl(2)).data(0).data0.in_point(i%)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).poi(1) = _
       m_lin(tl(2)).data(0).data0.in_point(i% + 1)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).para = "1"
Next i%
For i% = tn(3, 0) To tn(3, 1) - 1
con_length_of_polygon(last_conclusion).data(0).last_segment = _
     con_length_of_polygon(last_conclusion).data(0).last_segment + 1
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).line_no = tl(3)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).n(0) = i%
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).n(1) = i% + 1
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).poi(0) = _
       m_lin(tl(3)).data(0).data0.in_point(i%)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).poi(1) = _
       m_lin(tl(3)).data(0).data0.in_point(i% + 1)
con_length_of_polygon(last_conclusion).data(0). _
     segment(con_length_of_polygon(last_conclusion).data(0).last_segment).para = "1"
Next i%
   con_length_of_polygon(last_conclusion).record_.conclusion_no = last_conclusion + 1
   conclusion_data(last_conclusion).ty = length_of_polygon_
   ge_reduce_level = 3
temp_record.record_.conclusion_no = last_conclusion + 1
Call set_length_of_polygon(con_length_of_polygon(last_conclusion), 0, _
       temp_record)
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
End Sub
Public Sub draw_picture67(num As Integer)
Dim i%, j%
Dim ty(1) As Boolean
Dim ch As String * 1
Dim s(1) As String
Dim ts(1) As String
Dim tn(1) As Integer
Dim tn_(1) As Integer
Dim temp_record As total_record_type
Dim para(2) As String
'
Call string_from_wenti_condition(num, s)
'***********
For i% = 1 To Len(s(0))
ch = Mid$(s(0), i%, 1)
If ch >= "A" And ch <= "Z" Then
 ty(0) = True
  GoTo draw_picture67_mark1
End If
Next i%
draw_picture67_mark1:
For i% = 1 To Len(s(1))
ch = Mid$(s(1), i%, 1)
If ch >= "A" And ch <= "Z" Then
 ty(1) = True
  GoTo draw_picture67_mark2
End If
Next i%
draw_picture67_mark2:
If ty(0) = False And ty(1) = False Then
con_two_order_equation(last_conclusion).data(0).roots(0) = s(0)
con_two_order_equation(last_conclusion).data(0).roots(1) = s(1)
Else
Erase temp_item0
last_temp_item0 = 0
tn(0) = from_string_to_temp_item(s(0), ts(0))
tn(1) = from_string_to_temp_item(s(1), ts(1))
Call set_temp_item0(tn(0), -7, tn(1), -7, "*", "", "", s(0), tn_(1))
tn_(1) = simple_temp_item(tn_(0))
tn(0) = simple_temp_item(tn(0))
tn(1) = simple_temp_item(tn(1))
End If
If con_two_order_equation(last_conclusion).data(0).roots(0) <> "" And _
     con_two_order_equation(last_conclusion).data(0).roots(1) <> "" Then
   con_two_order_equation(last_conclusion).data(0).para(0) = "1"
   con_two_order_equation(last_conclusion).data(0).para(1) = time_string("-1", _
      add_string(con_two_order_equation(last_conclusion).data(0).roots(0), _
       con_two_order_equation(last_conclusion).data(0).roots(1), False, False), True, False)
   con_two_order_equation(last_conclusion).data(0).para(2) = time_string( _
     con_two_order_equation(last_conclusion).data(0).roots(0), _
      con_two_order_equation(last_conclusion).data(0).roots(1), True, False)
Else
  Call set_general_string(tn(0), tn(1), 0, 0, "-1", "-1", "0", "0", _
         "", last_conclusion, 0, 0, temp_record, tn_(0), 0)
  'Call set_item0(tn(0), -7, tn(1), -7, "*", 0, 0, 0, 0, 0, 0, _
        "", "", "1", "", "", last_conclusion, temp_record.record_data.data0.condition_data, _
          0, tn(1), 0)
  Call set_general_string(tn_(1), 0, 0, 0, "-1", "0", "0", "0", _
         "", last_conclusion, 0, 0, temp_record, tn_(1), 0)
  con_two_order_equation(last_conclusion).data(0).record.data0.condition_data.condition_no = 2
  con_two_order_equation(last_conclusion).data(0).record.data0.condition_data.condition(1).ty = general_string_
  con_two_order_equation(last_conclusion).data(0).record.data0.condition_data.condition(1).no = tn_(0)
  con_two_order_equation(last_conclusion).data(0).record.data0.condition_data.condition(2).ty = general_string_
  con_two_order_equation(last_conclusion).data(0).record.data0.condition_data.condition(2).no = tn_(1)
End If
conclusion_data(last_conclusion).ty = two_order_equation_
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True

End Sub

Public Sub draw_picture63(num As Integer)
'63 △□□□的面积=!_~
Dim i%
Dim value1 As String
Dim value2 As String
Dim triA%
Dim temp_record As total_record_type
i% = 3
While Asc(C_display_wenti.m_condition(num, i%)) > 13
 If C_display_wenti.m_condition(num, i%) < "A" Then
value1 = value1 + C_display_wenti.m_condition(num, i%)
i% = i% + 1
 Else
 value2 = value2 + C_display_wenti.m_condition(num, i%)
 i% = i% + 1
End If
Wend
If value2 <> "" And value1 <> "" Then
 value1 = value1 + "*" + value2
ElseIf value1 = "" Then
 value1 = value2
End If
triA% = triangle_number(C_display_wenti.m_point_no(num, 0), _
    C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
       0, 0, 0, 0, 0, 0, 0)
If set_or_prove < 2 Then
con_Area_of_element(last_conclusion).data(0).element.no = triA%
con_Area_of_element(last_conclusion).data(0).element.ty = triangle_
con_Area_of_element(last_conclusion).data(0).value = value1
If is_area_of_element(triangle_, triA%, _
        conclusion_data(last_conclusion).no(0), -1000) = False Then
         conclusion_data(last_conclusion).no(0) = 0
End If
  conclusion_data(last_conclusion).ty = area_of_element_
   area_of_triangle_conclusion = 1
 conclusion_data(last_conclusion).wenti_no = num
 last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
  MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
Call set_area_of_triangle(triA%, value1, temp_record, 0, 3)
area_of_triangle_conclusion = 1
End If
th_chose(155).chose = 1
th_chose(156).chose = 1
End Sub

Public Sub draw_picture62(num As Integer)
'62 △□□□的周长=!_~
Dim i%
Dim value1 As String
Dim value2 As String
Dim temp_record As total_record_type
i% = 3
While Asc(C_display_wenti.m_condition(num, i%)) > 13
 If C_display_wenti.m_condition(num, i%) < "A" And _
      C_display_wenti.m_condition(num, i%) > "Z" Then
  value1 = value1 + C_display_wenti.m_condition(num, i%)
   i% = i% + 1
 Else
 value2 = value2 + C_display_wenti.m_condition(num, i%)
 i% = i% + 1
End If
Wend
If value2 <> "" And value1 <> "" Then
 value1 = value1 + "*" + value2
ElseIf value1 = "" Then
 value1 = value2
End If
value1 = value_string(value1)
If set_or_prove < 2 Then
record_0.data0.condition_data.condition_no = 0 'record0
Call is_three_line_value(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 1), _
   C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 2), _
    C_display_wenti.m_point_no(num, 0), 0, 0, 0, 0, 0, 0, 0, 0, 0, _
     "1", "1", "1", value1, 0, -1000, 0, 0, 0, 0, 0, _
       con_line3_value(last_conclusion).data(0), 0, record_0.data0.condition_data, 0)
  conclusion_data(last_conclusion).ty = line3_value_
  line3_value_conclusion = 1
 conclusion_data(last_conclusion).wenti_no = num
 last_conclusion = last_conclusion + 1
'operate_step(num + 1).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(num + 1).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(num + 1).last_conclusion = last_conclusion
  MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Else
'record_0 = record0
Call set_three_line_value(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
  C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
   C_display_wenti.m_point_no(num, 4), C_display_wenti.m_point_no(num, 5), _
    0, 0, 0, 0, 0, 0, 0, 0, 0, "1", "1", "1", value1, temp_record, 0, 0, 0)
End If
End Sub
Public Sub draw_picture_56(num%, ByVal no_reduce As Byte)
Dim i%, j%
Dim it(3) As Integer
Dim v$
For i% = 0 To 3
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
     C_display_wenti.m_condition(num, i%)) Then
     Exit Sub
End If
   'If poi(C_display_wenti.m_point_no(num,i%)).data(0).degree > 2 Then
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   'End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
End Sub
Public Sub draw_equal_angle(ByVal p1%, ByVal p2%, ByVal p3%, p4%) '画角<p1%p2%,l3%平分线p2%p4%
Dim r(4) As Long
Dim A!
Dim A_(1) As Long
Dim t_coord(1) As POINTAPI
Dim sp_line%
If m_poi(4).data(0).parent.co_degree = 1 Or _
    (m_poi(4).data(0).parent.co_degree = 2 And m_poi(4).data(0).parent.element(2).ty <> line_) Then
 sp_line% = m_poi(4).data(0).parent.element(1).no
Else
 sp_line% = m_poi(4).data(0).parent.element(2).no
End If
Call line_sp_angle(p1%, p2%, p3%, sp_line%)
Call change_m_line(sp_line%)
'   r(3) = distance_of_two_POINTAPI(m_poi(p2%).data(0).data0.coordinate, m_poi(p4%).data(0).data0.coordinate)
'   r(4) = distance_of_two_POINTAPI(m_poi(p2%).data(0).data0.coordinate, t_coord(0))
'   A! = r(3) / r(4)
'   m_poi(p4%).data(0).data0.coordinate = add_POINTAPI(m_poi(p2%).data(0).data0.coordinate, _
           time_POINTAPI_by_number(minus_POINTAPI(t_coord(0), _
              m_poi(p2%).data(0).data0.coordinate), A!))
   'm_lin(m_poi(p4%).data(0).parent.element(1).no).data(0).data0.end_point_coord(1) = _
         m_poi(p4%).data(0).data0.coordinate
'    If m_poi(p4%).data(0).parent.inter_type = interset_point_line_line Then
'         m_lin(m_poi(p4%).data(0).parent.element(2).no).data(0).is_change = True
'          Call change_m_line(m_poi(p4%).data(0).parent.element(2).no, True)
'    Else
'      m_lin(m_poi(p4%).data(0).parent.element(1).no).data(0).is_change = True
'       Call change_m_line(m_poi(p4%).data(0).parent.element(1).no, True)
'    End If
'A! = r(1) / distance_of_two_POINTAPI(m_poi(p0%).data(0).data0.coordinate, _
                            m_poi(p4%).data(0).data0.coordinate)
't_coord1 = add_POINTAPI(m_poi(p0%).data(0).data0.coordinate, _
           time_POINTAPI_by_number(minus_POINTAPI( _
             m_poi(p4%).data(0).data0.coordinate, m_poi(p0%).data(0).data0.coordinate), A!))
'Call inter_point_circle_circle_(m_poi(p0%).data(0).data0.coordinate, r(2), _
                                 t_coord(0), r(0), _
                                  t_coord1, 0, t_coord2, 0, 0, 0, True)
'If ty = 0 Then
'A_(0) = cross_time_POINTAPI(minus_POINTAPI(m_poi(p1%).data(0).data0.coordinate, _
                             m_poi(p2%).data(0).data0.coordinate), _
                              minus_POINTAPI(m_poi(p3%).data(0).data0.coordinate, _
                               m_poi(p2%).data(0).data0.coordinate))
'A_(1) = cross_time_POINTAPI(minus_POINTAPI(t_coord1, _
                             m_poi(p0%).data(0).data0.coordinate), _
                              minus_POINTAPI(m_poi(p4%).data(0).data0.coordinate, _
                               m_poi(p0%).data(0).data0.coordinate))
'If (A_(0) > 0 And A_(1) > 0) Or (A_(0) < 0 And A_(1) < 0) Then
'    ty = 1
'Else
'    ty = 2
'End If
'End If
'If ty = 1 Then
' p_coord1 = t_coord1
'Else
' p_coord1 = t_coord2
'End If
'If l_no% > 0 Then
'   Call calculate_line_line_intersect_point(m_poi(p2%).data(0).data0.coordinate, p_coord1, _
            m_poi(m_lin(l_no%).data(0).data0.poi(0)).data(0).data0.coordinate, _
             m_poi(m_lin(l_no%).data(0).data0.poi(1)).data(0).data0.coordinate, _
              p_coord1, True)
'ElseIf C_no% > 0 Then
'   t_coord(0) = p_coord1
'   Call inter_point_line_circle3(m_poi(p2%).data(0).data0.coordinate, True, _
                   m_poi(p2%).data(0).data0.coordinate, t_coord(0), _
                    m_Circ(C_no%).data(0).data0, p_coord1, 0, p_coord2, 0, 0, True)
'End If
End Sub
Public Function line_sp_angle(ByVal p1%, ByVal p2%, ByVal p3%, Optional ByVal sp_line%) As Integer
Dim r1, r2 As Long
Dim r As Single
Dim p_coord As POINTAPI
'计算两侧边的长
r1 = distance_of_two_POINTAPI(m_poi(p1%).data(0).data0.coordinate, m_poi(p2%).data(0).data0.coordinate)
r2 = distance_of_two_POINTAPI(m_poi(p3%).data(0).data0.coordinate, m_poi(p2%).data(0).data0.coordinate)
'应用三角形顶角平分线的性质,计算顶角平分线与底边的交点的比
If r1 = 0 Or r2 = 0 Then
   Exit Function
End If
r = r1 / (r1 + r2) '按比例计算角平分线与底边交点的坐标
p_coord = add_POINTAPI(m_poi(p1%).data(0).data0.coordinate, _
           time_POINTAPI_by_number(minus_POINTAPI _
             (m_poi(p3%).data(0).data0.coordinate, m_poi(p1%).data(0).data0.coordinate), r))
             If sp_line% > 0 Then
              m_lin(sp_line%).data(0).data0.depend_poi(0) = p2%
              m_lin(sp_line%).data(0).data0.depend_poi1_coord = p_coord
              'm_lin(sp_line%).data(0).is_change = True
              'Call change_m_line(sp_line%)
              line_sp_angle = sp_line%
             Else
                If line_sp_angle = 0 Then
                     line_sp_angle = line_number(p2%, 0, m_poi(p2%).data(0).data0.coordinate, p_coord, _
                               depend_condition(point_, p2%), _
                               depend_condition(0, 0), _
                               aid_condition, fill_color, 0, 0)
               End If
             End If
End Function

Public Function inter_point_line_sp_angle_with_line(ByVal p1%, _
     ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, coord As POINTAPI, sp_line%, is_change As Boolean) As Boolean
If line_sp_angle(p1%, p2%, p3%, sp_line%) > 0 Then
   If line_number0(p1%, p3%, 0, 0) = line_number0(p4%, p5%, 0, 0) Then
      coord = t_coord
   Else
   inter_point_line_sp_angle_with_line = _
     calculate_line_line_intersect_point(m_poi(p2%).data(0).data0.coordinate, t_coord, m_poi(p4%).data(0).data0.coordinate, _
        m_poi(p5%).data(0).data0.coordinate, coord, is_change)
   End If
inter_point_line_sp_angle_with_line = True
End If
End Function

Public Function inter_point_line_sp_angle_with_circle(ByVal p1%, ByVal p2%, _
   ByVal p3%, c%, out_coord1 As POINTAPI, out_p1%, out_coord2 As POINTAPI, out_p2%)
If line_sp_angle(p1%, p2%, p3%) > 0 Then
  Call inter_point_line_circle3(m_poi(p2%).data(0).data0.coordinate, _
     True, m_poi(p2%).data(0).data0.coordinate, t_coord, m_Circ(c%).data(0).data0, _
       out_coord1, out_p1%, out_coord2, out_p2%, 0, False)
 inter_point_line_sp_angle_with_circle = True
Else
inter_point_line_sp_angle_with_circle = False
End If

End Function

Public Sub draw_picture_50(num%, ByVal no_reduce As Byte)
Dim i%
Dim t_ele1 As condition_type
Dim t_ele2 As condition_type
Dim coord As POINTAPI
Dim A!
For i% = 0 To 4
If C_display_wenti.m_point_no(num, i%) = 0 And i% <> 1 Then
Call draw_p0(num%, i%)
End If
Next i%
event_statue = 0
Exit Sub
Call line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 2), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                 condition, condition_color, 1, 0)
Call line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 3), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                 condition, condition_color, 1, 0)

If C_display_wenti.m_point_no(num, 1) = 0 Then
'********
If line_sp_angle(C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 0), _
      C_display_wenti.m_point_no(num, 3)) > 0 Then
 '     last_set_point = last_set_point + 1
 '      temp_point(6) = 100 - last_set_point
 '      Call set_set_point(temp_point(6))
       temp_point(6).no = 0
       Call set_aid_point(temp_point(6).no, coord, 1)
'End If
   Call C_display_wenti.set_m_point_no(num, _
      line_number0(C_display_wenti.m_point_no(num, 0), _
           temp_point(6).no, 0, 0), 5, False)
     'lin(C_display_wenti.m_point_no(num,5)).data(0).data0.visible = 4
draw_picture_50_mark4:
 event_statue = wait_for_draw_point
   While event_statue = wait_for_draw_point
     DoEvents
   Wend
 If event_statue = draw_point_down Or _
       event_statue = draw_point_move Or _
           event_statue = draw_point_up Then
   t_coord = input_coord
   ' temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture_50_mark4
      End If
     input_point_type% = read_inter_point(t_coord, t_ele1, _
                                t_ele2, temp_point(0).no, True)
      Call set_point_no_reduce(temp_point(0).no, 0)
         If input_point_type% <> new_point_on_line Or _
             t_ele1.no <> C_display_wenti.m_point_no(num, 5) Then
          If input_point_type% <> exist_point Then '不是旧的自由点
           Call remove_point(temp_point(0).no, display, 0)
          Else 'End If
           GoTo draw_picture_50_mark4
          End If
         End If
 
Call remove_point(temp_point(6).no, display, 0)
  Call C_display_wenti.set_m_point_no(num, temp_point(0).no, 1, False)
  Call set_point_name(C_display_wenti.m_point_no(num, 1), _
        C_display_wenti.m_condition(num, 1))
  m_poi(C_display_wenti.m_point_no(num, 1)).data(0).degree = 1
  Call set_point_in_line(C_display_wenti.m_point_no(num, 1), _
          C_display_wenti.m_point_no(num, 5))
  'Call put_name(C_display_wenti.m_point_no(num,1))
   '设置新点
   Call set_line_visible(C_display_wenti.m_point_no(num, 5), 1)
    temp_line(1) = line_number(temp_point(0).no, C_display_wenti.m_point_no(num, 0), _
                               pointapi0, pointapi0, _
                               depend_condition(point_, temp_point(0).no), _
                               depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                               condition, condition_color, 1, 0)
 End If
          A! = distance_of_two_POINTAPI(m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate, _
                 m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate) / _
               (distance_of_two_POINTAPI(m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate, _
                 m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate) + _
                distance_of_two_POINTAPI(m_poi(C_display_wenti.m_point_no(num, 3)).data(0).data0.coordinate, _
             m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate))
          Call C_display_wenti.set_m_point_no(num, CInt(1000 * A!), 7, False)
End If
End Sub

Public Sub draw_picture_51(num%, ByVal no_reduce As Byte)
Dim i%
Dim coord As POINTAPI
For i% = 0 To 4
 If C_display_wenti.m_point_no(num, i%) = 0 Then
  Call draw_p0(num%, i%)
 End If
Next i%
Call line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                 condition, condition_color, 1, 0)
Call line_number(C_display_wenti.m_point_no(num, 2), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                 condition, condition_color, 1, 0)
Call line_number(C_display_wenti.m_point_no(num, 3), _
                 C_display_wenti.m_point_no(num, 4), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 4)), _
                 condition, condition_color, 1, 0)
If C_display_wenti.m_point_no(num, 5) = 0 Then
If inter_point_line_sp_angle_with_line(C_display_wenti.m_point_no(num, 0), _
  C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
    C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 4), _
     coord, False) Then
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
MDIForm1.Toolbar1.Buttons(21).Image = 33
  ' Call init_Point0(last_conditions.last_cond(1).point_no)
Call C_display_wenti.set_m_point_no(num, last_conditions.last_cond(1).point_no, 5, True)
 Call set_point_coordinate(C_display_wenti.m_point_no(num, 5), coord, False)
 Call set_point_name(C_display_wenti.m_point_no(num, 5), _
                                 C_display_wenti.m_condition(num, 5))
 Call set_point_visible(C_display_wenti.m_point_no(num, 5), 1, False)
 'Call draw_point(Draw_form, m_poi(C_display_wenti.m_point_no(num,5)), 0, display)
 record_0.data0.condition_data.condition_no = 0
 Call add_point_to_line(C_display_wenti.m_point_no(num, 5), _
    line_number(C_display_wenti.m_point_no(num, 3), _
                C_display_wenti.m_point_no(num, 4), _
                pointapi0, pointapi0, _
                depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                depend_condition(point_, C_display_wenti.m_point_no(num, 4)), _
                condition, condition_color, 1, 0), 0, display, True, 0, temp_record)
End If
End If
End Sub

Public Sub draw_picture_52(num As Integer, ByVal no_reduce As Byte)
Dim i%, c%
Dim t_ele1 As condition_type
Dim t_ele2 As condition_type
'c_display_wenti.no = -52
For i% = 0 To 4
  If C_display_wenti.m_point_no(num, i%) = 0 Then
    Call draw_p0(num%, i%)
  End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
Call line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                 condition, condition_color, 1, 0)
Call line_number(C_display_wenti.m_point_no(num, 2), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                 condition, condition_color, 1, 0)
c% = m_circle_number(1, C_display_wenti.m_point_no(num, 3), pointapi0, _
               C_display_wenti.m_point_no(num, 4), 0, 0, 0, 0, 0, _
                1, 1, condition, condition_color, True)
Call C_display_wenti.set_m_point_no(num, c%, 9, False)
If C_display_wenti.m_point_no(num, 5) = 0 Then
If inter_point_line_sp_angle_with_circle(C_display_wenti.m_point_no(num, 0), _
     C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
      c%, t_coord1, 0, t_coord2, 0) Then
'       last_set_point = last_set_point + 1
'        temp_point(6) = 100 - last_set_point
'         Call set_set_point(temp_point(6))
t_coord1.X = _
           m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X + 10 * _
             (t_coord1.X - m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X)
t_coord1.Y = _
           m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y + 10 * _
             (t_coord1.Y - m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y)
     temp_point(6).no = 0
     Call set_aid_point(temp_point(6).no, t_coord1, 1)
    temp_line(0) = line_number0(C_display_wenti.m_point_no(num, 1), _
                 temp_point(6).no, 0, 0)
   'lin(temp_line(0)).data(0).data0.visible = 4
draw_picture_52_mark4:
 event_statue = wait_for_draw_point
   While event_statue = wait_for_draw_point
     DoEvents
      Wend
 If event_statue = draw_point_down Or _
       event_statue = draw_point_move Or _
           event_statue = draw_point_up Then
   t_coord = input_coord
   ' temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture_52_mark4
      End If
     input_point_type% = read_inter_point(t_coord, t_ele1, _
                                    t_ele2, temp_point(0).no, True)
     Call set_point_no_reduce(temp_point(0).no, 0)
      If (input_point_type% = new_point_on_line_circle12 Or _
       input_point_type% = new_point_on_line_circle21 Or _
        input_point_type% = new_point_on_line_circle) And _
   t_ele1.no = temp_line(0) And t_ele2.no = c% Then
 If input_point_type% = new_point_on_line_circle21 Then
    Call C_display_wenti.set_m_point_no(num, 1, 10, True)
 End If
  Call C_display_wenti.set_m_point_no(num, temp_point(0).no, 5, True)
   Call remove_point(temp_point(1).no, display, 0)
   'poi(C_display_wenti.m_point_no(num,5)).data(0).data0.name = C_display_wenti.m_condition(5)
    'Call put_name(C_display_wenti.m_point_no(num,5))
     Call remove_point(temp_point(6).no, display, 0)
  Call set_point_in_line(C_display_wenti.m_point_no(num, 5), temp_line(0))
   '设置新点
   Call set_line_visible(temp_line(0), 1)
    temp_line(0) = line_number(C_display_wenti.m_point_no(num, 5), _
                               C_display_wenti.m_point_no(num, 1), _
                               pointapi0, pointapi0, _
                               depend_condition(point_, C_display_wenti.m_point_no(num, 5)), _
                               depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                               condition, condition_color, 1, 0)
End If
Else
               Call remove_point(temp_point(0).no, display, 0)
                Call remove_point(temp_point(1).no, display, 0)
                 GoTo draw_picture_52_mark4
   End If
End If
  Call add_point_to_m_circle( _
            C_display_wenti.m_point_no(num, 5), c%, record0, record0, 255)
'******************
End Sub

Public Sub draw_picture_43_42(num As Integer, in_wenti_data As wentitype, ByVal no_reduce As Byte)
Dim i%, j%, k%
Dim ele1 As condition_type
Dim ele2 As condition_type
Dim value As String
Dim A!
For i% = 0 To 3
If draw_free_point(in_wenti_data.point_no(i%), _
                      in_wenti_data.condition(i%)) Then
    Exit Sub
End If
Next i%
   'If i% <> 0 Then
   ' Call change_point_degree(C_display_wenti.m_point_no(num, 3), -3)
   'End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
If in_wenti_data.no = -42 Then
  If Ratio_for_measure.Ratio_for_measure = 0 Then '第一次输入长度
    Ratio_for_measure.Ratio_for_measure = m_Circ(in_wenti_data.circ(1)).data(0).data0.radii / _
     m_Circ(in_wenti_data.circ(1)).data(0).data0.real_radii  '设置屏幕长度和实际长度的比
      Call set_son_data(circle_, in_wenti_data.circ(1), Ratio_for_measure.sons) '设置继承关系
   Else
     m_Circ(in_wenti_data.circ(1)).data(0).data0.radii = m_Circ(in_wenti_data.circ(1)).data(0).data0.real_radii * _
     Ratio_for_measure.Ratio_for_measure '显示的半径
      m_Circ(in_wenti_data.circ(1)).data(0).is_change = True
       '计算交点
        Call change_m_circle(in_wenti_data.circ(1), depend_condition(0, 0))
   End If
'Call C_display_wenti.set_m_point_no(num, m_circle_number( _
    1, C_display_wenti.m_point_no(num, 0), pointapi0, _
     C_display_wenti.m_point_no(num, 1), 0, 0, 0, 0, 0, _
      1, 1, condition, condition_color, True), 11, False)
Else
'Call C_display_wenti.set_m_point_no(num, line_number( _
      C_display_wenti.m_point_no(num, 0), _
      C_display_wenti.m_point_no(num, 1), _
      pointapi0, pointapi0, _
      depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
      depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
      condition, condition_color, 1, 0), 11, False)
End If
draw_picture1_mark_42:
 event_statue = wait_for_draw_point
   While event_statue = wait_for_draw_point
     DoEvents
      Wend
  If event_statue = draw_point_down Or _
     event_statue = draw_point_move Or _
        event_statue = draw_point_up Then 'mouse_type <> 1 Then
   t_coord = input_coord
   ' temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark_42
 End If
      input_point_type% = read_inter_point(t_coord, ele1, _
                            ele2, temp_point(1).no, True)
      Call set_point_no_reduce(temp_point(1).no, 0)
If ((input_point_type% = new_point_on_circle And _
      C_display_wenti.m_no(num) = -42) Or ( _
    input_point_type% = new_point_on_line And _
      C_display_wenti.m_no(num) = -43)) And _
         ele1.no = C_display_wenti.m_point_no(num, 11) Then
 Call C_display_wenti.set_m_point_no(num, temp_point(1).no, 2, True)
  'poi(C_display_wenti.m_point_no(num,2)).data(0).data0.name = _
      C_display_wenti.m_condition (2)
'Call put_name(C_display_wenti.m_point_no(num,2))
Else
 If input_point_type% <> exist_point Then
  Call remove_point(temp_point(2).no, display, 0)
 End If
 GoTo draw_picture1_mark_42
End If
End Sub


Public Sub draw_picture_40(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%
Dim tp(7) As Integer
Dim tl(3) As Integer
For i% = 0 To 7
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
   C_display_wenti.m_condition(num, i%)) Then
Exit Sub
End If
   'If poi(C_display_wenti.m_point_no(num,i%)).data(0).degree > 2 Then
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   'End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
End Sub

Public Sub draw_picture_39(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%
For i% = 0 To 8
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
   C_display_wenti.m_condition(num, i%)) Then
Exit Sub
End If
   'If poi(C_display_wenti.m_point_no(num,i%)).data(0).degree > 2 Then
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   'End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
End Sub

Public Sub draw_picture_38(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%
Dim value1 As String
Dim ang(1) As Integer
For i% = 0 To 5
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
      C_display_wenti.m_condition(num, i%)) Then
      Exit Sub
End If
   'If poi(C_display_wenti.m_point_no(num,i%)).data(0).degree > 2 Then
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   'End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
End Sub

Public Sub draw_picture_37(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%
For i% = 0 To 5
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
      C_display_wenti.m_condition(num, i%)) Then
  Exit Sub
End If
   'If poi(C_display_wenti.m_point_no(num,i%)).data(0).degree > 2 Then
   Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   'End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
End Sub

Public Sub draw_picture_36(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%
For i% = 0 To 5
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
     C_display_wenti.m_condition(num, i%)) Then
  Exit Sub
End If
   'If poi(C_display_wenti.m_point_no(num,i%)).data(0).degree > 2 Then
   Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   'End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
End Sub
Public Sub draw_picture_35(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%
For i% = 0 To 5
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
    C_display_wenti.m_condition(num, i%)) Then
    Exit Sub
End If
   'If poi(C_display_wenti.m_point_no(num,i%)).data(0).degree > 2 Then
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   'End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
End Sub

Public Sub draw_picture_34(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%
Dim value1 As String
For i% = 0 To 3
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
    C_display_wenti.m_condition(num, i%)) Then
    Exit Sub
End If
'   If poi(C_display_wenti.m_point_no(num,i%)).data(0).degree > 2 Then
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
'   End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
End Sub

Public Sub draw_picture_24(ByVal num As Integer, ByVal no_reduce As Byte)
'-24 弧□□＝弧□□
Dim i%, k%, m%, n%
For i% = 0 To 3
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
     C_display_wenti.m_condition(num, i%)) Then
     Exit Sub
End If
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
Next i%
End Sub
Public Sub draw_picture_23_22(ByVal num As Integer, ByVal no_reduce As Byte)
'-22 过□点平行□□的直线交□□于□
'-23过□点垂直□□的直线交□□于□
Dim i%
Dim t_cond As condition_data_type
Dim ty(1) As Boolean
For i% = 0 To 4
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
   C_display_wenti.m_condition(num, i%)) Then
   Exit Sub
End If
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
Next i%
temp_line(0) = line_number(C_display_wenti.m_point_no(num, 1), _
                           C_display_wenti.m_point_no(num, 2), _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                           condition, condition_color, 1, 0)
temp_line(1) = line_number(C_display_wenti.m_point_no(num, 3), _
                           C_display_wenti.m_point_no(num, 4), _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 4)), _
                           condition, condition_color, 1, 0)
If C_display_wenti.m_no(num) = -22 Then
  ty(0) = True
Else
  ty(0) = False
End If
 last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
 MDIForm1.Toolbar1.Buttons(21).Image = 33
  ' Call init_Point0(last_conditions.last_cond(1).point_no)
     Call C_display_wenti.set_m_point_no(num, last_conditions.last_cond(1).point_no, 5, False)
Call inter_point_line_line3(C_display_wenti.m_point_no(num, 0), _
  ty(0), line_number0(C_display_wenti.m_point_no(num, 1), _
      C_display_wenti.m_point_no(num, 2), 0, 0), _
         C_display_wenti.m_point_no(num, 3), _
   True, line_number0(C_display_wenti.m_point_no(num, 4), _
            C_display_wenti.m_point_no(num, 3), 0, 0), _
             t_coord, C_display_wenti.m_point_no(num, 5), False, t_cond, True)
temp_line(2) = line_number(C_display_wenti.m_point_no(num, 0), _
                           C_display_wenti.m_point_no(num, 5), _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 5)), _
                           condition, condition_color, 1, 0)
m_poi(C_display_wenti.m_point_no(num, 5)).data(0).degree = 1
Call set_point_name(C_display_wenti.m_point_no(num, 5), _
          C_display_wenti.m_condition(num, 5))
m_poi(C_display_wenti.m_point_no(num, 5)).data(0).degree = 0
record_0.data0.condition_data.condition_no = 0
 Call add_point_to_line( _
    C_display_wenti.m_point_no(num, 5), temp_line(1), 0, display, _
     True, 0, temp_record)
'Call draw_point(Draw_form, m_poi(C_display_wenti.m_point_no(num,5)), 0, display)
If C_display_wenti.m_no(num) = -22 Then
      Call paral_line(temp_line(0), temp_line(2), True, True) '
Else
      Call vertical_line(temp_line(0), temp_line(2), True, True) '
End If
End Sub

Public Sub draw_picture_20(ByVal num As Integer, ByVal no_reduce As Byte)
'-20 任意△□□□
Call C_display_wenti.Get_wenti(num)
Call draw_any_triangle(wenti_cond0.data)
End Sub

Public Sub draw_picture_21(ByVal num As Integer, ByVal no_reduce As Byte)
'-21 △□□□是直角三角形
Dim i%
Dim t_ele1 As condition_type
Dim t_ele2 As condition_type
Dim A!
For i% = 0 To 1
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
     C_display_wenti.m_condition(num, i%)) Then
  Exit Sub
End If
   Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
Next i%
If C_display_wenti.m_point_no(num, i%) = 0 Then
 '垂线上任取一点
'******************************************************
 '    last_set_point = last_set_point + 1
 '      temp_point(6) = 100 - last_set_point
 '      Call set_set_point(temp_point(6))
     t_coord.X = _
      m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X + _
       m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y - _
        m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.Y
     t_coord.Y = _
      m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y - _
       m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X + _
        m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.X
        temp_point(6).no = 0
        Call set_aid_point(temp_point(6).no, t_coord, 1)
     temp_line(1) = line_number0(C_display_wenti.m_point_no(num, 0), _
           temp_point(6).no, 0, 0)
      'lin(temp_line(1)).data(0).data0.visible = 4
draw_picture_21_mark2:
    event_statue = wait_for_draw_point
     While event_statue = wait_for_draw_point
     DoEvents
      Wend
 If event_statue = draw_point_down Or _
    event_statue = draw_point_move Or _
        event_statue = draw_point_up Then 'mouse_type <> 1 Then
   t_coord = input_coord
   ' temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture_21_mark2
 End If
     input_point_type% = read_inter_point(t_coord, t_ele1, _
                                       t_ele2, temp_point(0).no, True)
         Call set_point_no_reduce(temp_point(0).no, 0)
         If input_point_type% <> new_point_on_line Or _
             t_ele1.no <> temp_line(1) Then
         If input_point_type% <> exist_point Then  '不是旧的自由点
          Call remove_point(temp_point(0).no, display, 0)
          GoTo draw_picture_21_mark2
          Else 'End If
          GoTo draw_picture_21_mark2
          End If
         End If
      Call draw_tangent_line(0)
      Call draw_tangent_line(1)
      Call remove_point(temp_point(6).no, display, 0)
      temp_line(1) = line_number(C_display_wenti.m_point_no(num, 0), _
                                 temp_point(0).no, pointapi0, pointapi0, _
                                 depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                                 depend_condition(point_, temp_point(0).no), _
                                 condition, condition_color, 1, 0)
      Call set_point_name(temp_point(0).no, C_display_wenti.m_condition(num, 2))
      m_poi(temp_point(0).no).data(0).degree = 1
      Call set_point_in_line(temp_point(0).no, temp_line(1))
      Call C_display_wenti.set_m_point_no(num, temp_point(0).no, 3, True)
          If Abs(m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X - _
            m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X) > 4 Then
           A! = -(m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.Y - _
             m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y) / _
               (m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X - _
                m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X)
          Else
           A! = (m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.X - _
              m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X) / _
               (m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y - _
                m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y)
          End If
    'Call put_name(C_display_wenti.m_point_no(num,2))
  Call C_display_wenti.set_m_point_no(num, Int(A! * 1000), 7, False)
End If

  temp_line(0) = line_number(C_display_wenti.m_point_no(num, 1), _
                             C_display_wenti.m_point_no(num, 0), _
                             pointapi0, pointapi0, _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                           condition, condition_color, 1, 0)
  temp_line(1) = line_number(C_display_wenti.m_point_no(num, 0), _
                             C_display_wenti.m_point_no(num, 2), _
                             pointapi0, pointapi0, _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                             condition, condition_color, 1, 0)
  temp_line(2) = line_number(C_display_wenti.m_point_no(num, 1), _
                             C_display_wenti.m_point_no(num, 2), _
                             pointapi0, pointapi0, _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                             condition, condition_color, 1, 0)
Call vertical_line(temp_line(0), temp_line(1), True, True)
End Sub

Public Sub draw_picture_19(ByVal num As Integer, ByVal no_reduce As Byte)
'-19 任意四边形□□□□
Call C_display_wenti.Get_wenti(num)
Call draw_any_polygon4(wenti_cond0.data)
End Sub
Public Sub draw_picture_16_12_9_8(ByVal num As Integer, ByVal no_reduce As Byte)
'-8 □□□□□□是正六边形
'-9 □□□□□是正五边形
'-12 □□□□是正方形
'-16 △□□□是等边三角形
Dim i%, j%, poly_no%
Dim t_p As POINTAPI
Dim ty(1) As Boolean
Dim pol As polygon
 If C_display_wenti.m_no(num) = -16 Then
  j% = 3
 ElseIf C_display_wenti.m_no(num) = -12 Then
  j% = 4
 ElseIf C_display_wenti.m_no(num) = -9 Then
  j% = 5
 ElseIf C_display_wenti.m_no(num) = -8 Then
  j% = 6
 End If
For i% = 0 To 1
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
     C_display_wenti.m_condition(num, i%)) Then
     Exit Sub
End If
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
 Next i%
draw_picture1_mark_16:
  event_statue = wait_for_draw_point  '输点状态
   While event_statue = wait_for_draw_point '等待事件发生
    DoEvents
   Wend
 If event_statue = draw_point_down Then 'Or event_statue = _
             draw_point_move Or event_statue = _
                    draw_point_up Then 'mouse_type <> 1 Then
   t_p.X = input_coord.X
    t_p.Y = input_coord.Y
 Else
     GoTo draw_picture1_mark_16
 End If
 If area_triangle(m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate, _
     m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate, t_p) > 0 Then
 ty(0) = True
 Else
 ty(0) = False
 End If
poly_no% = 0
Call set_polygon(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
        j%, poly_no%, ty(0), True, num)
          Call C_display_wenti.set_m_point_no(num, _
                poly(poly_no%).v(2), 2, True)
          Call C_display_wenti.set_m_point_no(num, _
                poly_no%, 10, False)
End Sub

Public Sub draw_picture_13_11_14_10(ByVal num As Integer, ByVal no_reduce As Byte)
'-10 □□□□是菱形
'-11 □□□□是平行四边形
'-13 □□□□是长方形
'-14 □□□□是等腰梯形
Dim i%, t_line%
Dim dr_ty As Integer
For i% = 0 To 1
 If C_display_wenti.m_point_no(num, i%) = 0 Then
  Call draw_p0(num, i%)
 End If
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
Next i%
 t_line% = line_number(C_display_wenti.m_point_no(num, 0), _
                       C_display_wenti.m_point_no(num, 1), _
                       pointapi0, pointapi0, _
                       depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                       depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                       condition, condition_color, 1, 0)
If display_temp_four_point_fig = 0 Then
temp_four_point_fig.poi(0) = C_display_wenti.m_point_no(num, 0)
temp_four_point_fig.poi(1) = C_display_wenti.m_point_no(num, 1)
temp_four_point_fig.poi(2) = C_display_wenti.m_point_no(num, 2)
temp_four_point_fig.poi(3) = C_display_wenti.m_point_no(num, 3)
temp_four_point_fig.p(0) = _
     m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate
temp_four_point_fig.p(1) = _
     m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate
If temp_four_point_fig.poi(2) > 0 And temp_four_point_fig.poi(3) = 0 Then
   mouse_move_coord = m_poi(temp_four_point_fig.poi(2)).data(0).data0.coordinate
     Call C_display_wenti.set_m_point_no(num, 1, 47, False)
      dr_ty = 1
ElseIf temp_four_point_fig.poi(3) > 0 And temp_four_point_fig.poi(2) = 0 Then
   mouse_move_coord = m_poi(temp_four_point_fig.poi(3)).data(0).data0.coordinate
      Call C_display_wenti.set_m_point_no(num, 2, 47, False)
       dr_ty = 2
ElseIf temp_four_point_fig.poi(3) > 0 And temp_four_point_fig.poi(2) > 0 Then
    If m_poi(temp_four_point_fig.poi(3)).data(0).degree > _
        m_poi(temp_four_point_fig.poi(2)).data(0).degree Then
   mouse_move_coord = m_poi(temp_four_point_fig.poi(2)).data(0).data0.coordinate
     Call C_display_wenti.set_m_point_no(num, 1, 47, False)
      dr_ty = 1
    Else
   mouse_move_coord = m_poi(temp_four_point_fig.poi(3)).data(0).data0.coordinate
      Call C_display_wenti.set_m_point_no(num, 2, 47, False)
       dr_ty = 2
    End If
End If
'*****************************************************************
If C_display_wenti.m_no(num) = -13 Then
   Call draw_temp_long_squre(move_coord, dr_ty)
        display_temp_four_point_fig = 1
         temp_four_point_fig_type = long_squre_
ElseIf C_display_wenti.m_no(num) = -11 Then
  Call draw_temp_parallelogram(move_coord, dr_ty)
        display_temp_four_point_fig = 1
         temp_four_point_fig_type = parallelogram_
ElseIf C_display_wenti.m_no(num) = -14 Then
  Call draw_temp_equal_side_tixing(move_coord, dr_ty)
        display_temp_four_point_fig = 1
         temp_four_point_fig_type = equal_side_tixing_
ElseIf C_display_wenti.m_no(num) = -10 Then
  Call draw_temp_rhombus(move_coord, dr_ty)
       display_temp_four_point_fig = 1
        temp_four_point_fig_type = rhombus_
End If
End If
If temp_four_point_fig.poi(2) > 0 Or _
     temp_four_point_fig.poi(3) > 0 Then
 GoTo draw_picture_13_mark10
End If
draw_picture_13_mark0:
 event_statue = wait_for_draw_point  '输点状态
   While event_statue = wait_for_draw_point '等待事件发生
    DoEvents
   Wend
If event_statue = draw_point_down Then
 '???
Else
 GoTo draw_picture_13_mark0
End If
draw_picture_13_mark10:
If temp_four_point_fig.poi(2) > 0 Then
Call C_display_picture.set_m_point_coordinate(temp_four_point_fig.poi(2), _
        temp_four_point_fig.p(2).X, temp_four_point_fig.p(2).Y)
Else
temp_four_point_fig.poi(2) = m_point_number(temp_four_point_fig.p(2), condition, 1, condition_color, _
                      C_display_wenti.m_condition(num, 2), _
                         condition_type0, condition_type0, 0, True)
MDIForm1.Toolbar1.Buttons(21).Image = 33
'   Call init_Point0(last_conditions.last_cond(1).point_no)
 Call C_display_wenti.set_m_point_no(num, _
     last_conditions.last_cond(1).point_no, 2, False)
 m_poi(last_conditions.last_cond(1).point_no).data(0).degree = 2
End If
'*******************************
If temp_four_point_fig.poi(3) > 0 Then
Call C_display_picture.set_m_point_coordinate(temp_four_point_fig.poi(3), _
        temp_four_point_fig.p(3).X, temp_four_point_fig.p(3).Y)
Else
temp_four_point_fig.poi(3) = m_point_number(temp_four_point_fig.p(3), condition, 1, condition_color, _
                                           C_display_wenti.m_condition(num, 3), _
                                            condition_type0, condition_type0, 0, True)
                                           
MDIForm1.Toolbar1.Buttons(21).Image = 33
   Call C_display_wenti.set_m_point_no(num, last_conditions.last_cond(1).point_no, 3, False)  '重画用
End If
If C_display_wenti.m_no(num) = -13 Then
    Call draw_temp_long_squre(temp_four_point_fig.p(2), 1)
       display_temp_four_point_fig = 0
ElseIf C_display_wenti.m_no(num) = -11 Then
   Call draw_temp_parallelogram(temp_four_point_fig.p(2), 1)
       display_temp_four_point_fig = 0
ElseIf C_display_wenti.m_no(num) = -14 Then
  Call draw_temp_equal_side_tixing(temp_four_point_fig.p(2), 1)
       display_temp_four_point_fig = 0
ElseIf C_display_wenti.m_no(num) = -10 Then
  Call draw_temp_rhombus(temp_four_point_fig.p(2), 1)
       display_temp_four_point_fig = 0
End If
   Call line_number(C_display_wenti.m_point_no(num, 1), _
                    C_display_wenti.m_point_no(num, 2), _
                    pointapi0, pointapi0, _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                    condition, condition_color, 1, 0)
   Call line_number(C_display_wenti.m_point_no(num, 3), _
                    C_display_wenti.m_point_no(num, 2), _
                    pointapi0, pointapi0, _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                    condition, condition_color, 1, 0)
   Call line_number(C_display_wenti.m_point_no(num, 0), _
                    C_display_wenti.m_point_no(num, 3), _
                    pointapi0, pointapi0, _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                    condition, condition_color, 1, 0)
'*******
End Sub

Public Sub draw_picture_7(ByVal num As Integer, ByVal no_reduce As Byte)
'-7 □□/□□=!_~
Dim i%
Dim value1 As String
For i% = 0 To 3
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
      C_display_wenti.m_condition(i%, num)) Then
      Exit Sub
End If
   Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
Next i%
End Sub

Public Sub draw_picture_6_43_42_57(ByVal num As Integer, ByVal no_reduce As Byte)
'-6 □□=!_~
'6□是线段□□上分比为!_~的分点
Dim i%, l%
Dim value As String
Dim wenti_data As wentitype
For i% = 0 To 1
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
     C_display_wenti.m_condition(num, i%)) Then
     Exit Sub
End If
    'Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
Next i%
Call C_display_wenti.Get_wenti(num) '获取wenti_cond0即输入语句数据
wenti_data = wenti_cond0.data
If wenti_data.no = -6 Then
 value = number_string(C_display_wenti.m_point_no(num, 2))
ElseIf wenti_data.no = -43 Or wenti_data.no = -42 Then
 value = number_string(C_display_wenti.m_point_no(num, 4))
ElseIf wenti_data.no = -57 Then
 value = number_string(C_display_wenti.m_point_no(num, 5))
End If
      m_Circ(wenti_data.circ(1)).data(0).data0.real_radii = value_string(value) '圆的真实半径（输入条件）
      'C_display_wenti.m_point_no(num, 2) 输入数字在number_string中位置
If Ratio_for_measure.Ratio_for_measure = 0 Then '第一次输入长度
  Ratio_for_measure.Ratio_for_measure = m_Circ(wenti_data.circ(1)).data(0).data0.radii / _
     m_Circ(wenti_data.circ(1)).data(0).data0.real_radii  '设置屏幕长度和实际长度的比
Else '按设置屏幕长度和实际长度的比，重新设置圆的半径，
  m_Circ(wenti_data.circ(1)).data(0).data0.radii = m_Circ(wenti_data.circ(1)).data(0).data0.real_radii * _
     Ratio_for_measure.Ratio_for_measure '显示的半径
      m_Circ(wenti_data.circ(1)).data(0).is_change = True
        Call change_m_circle(wenti_data.circ(1), depend_condition(0, 0))
End If
 If m_poi(m_Circ(wenti_data.circ(1)).data(0).data0.center).data(0).parent.co_degree >= 2 And _
      m_poi(m_Circ(wenti_data.circ(1)).data(0).data0.in_point(1)).data(0).parent.co_degree >= 2 Then
  Ratio_for_measure.Ratio_for_measure = m_Circ(wenti_data.circ(1)).data(0).data0.radii / _
     m_Circ(wenti_data.circ(1)).data(0).data0.real_radii  '设置屏幕长度和实际长度的比
       Ratio_for_measure.is_fixed_ratio = True
        Call change_ratio_for_measure
  End If
      Call set_parent(Ratio_for_measure_, 1, circle_, wenti_data.circ(1), 0) '设置继承关系
      Ratio_for_measure.ratio_for_measure0 = Ratio_for_measure.Ratio_for_measure
 'End If
End Sub

Public Sub draw_picture2_3(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%, j%
Dim t_ele1 As condition_type
Dim t_ele2 As condition_type
Dim A!
If C_display_wenti.m_point_no(num, 1) = 0 Then
 j% = 1
ElseIf C_display_wenti.m_point_no(num, 3) = 0 Then
 j% = 3
ElseIf C_display_wenti.m_point_no(num, 0) = 0 Then
 j% = 0
ElseIf C_display_wenti.m_point_no(num, 2) = 0 Then
 j% = 2
Else
 j% = 4
End If
For i% = 0 To 3
 If j% <> i% Then
 If draw_free_point(C_display_wenti.m_point_no(num, i%), _
  C_display_wenti.m_condition(num, i%)) Then
   Exit Sub
 End If
 End If
Next i%
If j% < 4 Then '非限制性输入
 If j% = 0 Or j% = 1 Then
  temp_point(2).no = C_display_wenti.m_point_no(num, 2)
     temp_point(3).no = C_display_wenti.m_point_no(num, 3)
      temp_line(0) = line_number(C_display_wenti.m_point_no(num, 2), _
                                 C_display_wenti.m_point_no(num, 3), _
                                 pointapi0, pointapi0, _
                                 depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                                 depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                                 condition, condition_color, 1, 0) '(k%, l%)
 ElseIf j% = 2 Or j% = 3 Then
     temp_point(2).no = C_display_wenti.m_point_no(num, 0)
      temp_point(3).no = C_display_wenti.m_point_no(num, 1)
       temp_line(0) = line_number(C_display_wenti.m_point_no(num, 0), _
                                  C_display_wenti.m_point_no(num, 1), _
                                  pointapi0, pointapi0, _
                                  depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                                  depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                                  condition, condition_color, 1, 0)  '(k%, l%)
End If
 If j% = 0 Then
 temp_point(1).no = C_display_wenti.m_point_no(num, 1)
 ElseIf j% = 1 Then
 temp_point(1).no = C_display_wenti.m_point_no(num, 0)
 ElseIf j% = 2 Then
 temp_point(1).no = C_display_wenti.m_point_no(num, 3)
 ElseIf j% = 3 Then
 temp_point(1).no = C_display_wenti.m_point_no(num, 2)
 End If
'******************************************************
'     last_set_point = last_set_point + 1
'       temp_point(6) = 100 - last_set_point
'       Call set_set_point(temp_point(6))
     If C_display_wenti.m_no(num) = 2 Then
     t_coord.X = m_poi(temp_point(1).no).data(0).data0.coordinate.X + _
      m_poi(temp_point(2).no).data(0).data0.coordinate.X - m_poi(temp_point(3).no).data(0).data0.coordinate.X
     t_coord.Y = m_poi(temp_point(1).no).data(0).data0.coordinate.Y + _
      m_poi(temp_point(2).no).data(0).data0.coordinate.Y - m_poi(temp_point(3).no).data(0).data0.coordinate.Y
      temp_point(6).no = 0
      Call set_aid_point(temp_point(6).no, t_coord, 1)
    Else
     t_coord.X = m_poi(temp_point(1).no).data(0).data0.coordinate.X + _
      m_poi(temp_point(2).no).data(0).data0.coordinate.Y - m_poi(temp_point(3).no).data(0).data0.coordinate.Y
     t_coord.Y = m_poi(temp_point(1).no).data(0).data0.coordinate.Y - _
        m_poi(temp_point(2).no).data(0).data0.coordinate.X + m_poi(temp_point(3).no).data(0).data0.coordinate.X
      Call set_point_coordinate(temp_point(6).no, t_coord, False)
    End If
        
     If C_display_wenti.m_no(num) = 2 Then
      If is_point_in_line3(temp_point(1).no, m_lin(temp_line(0)).data(0).data0, 0) Then
         temp_line(1) = temp_line(0)
         ' Call add_point_to_line(temp_point(2), temp_line(1), True, display)
      Else
     temp_line(1) = line_number0(temp_point(1).no, temp_point(6).no, 0, 0)
      'lin(temp_line(1)).data(0).data0.visible = 4
       End If
     ElseIf C_display_wenti.m_no(num) = 3 Then
     temp_line(1) = line_number0(temp_point(1).no, temp_point(6).no, 0, 0)
      'lin(temp_line(1)).data(0).data0.visible = 4
      End If
draw_picture1_mark2:
    event_statue = wait_for_draw_point
     While event_statue = wait_for_draw_point
     DoEvents
      Wend
 If event_statue = draw_point_down Or _
    event_statue = draw_point_move Or _
        event_statue = draw_point_up Then 'mouse_type <> 1 Then
   t_coord = input_coord
   ' temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark2
 End If
     input_point_type% = read_inter_point(t_coord, t_ele1, _
                                   t_ele2, temp_point(0).no, True)
         Call set_point_no_reduce(temp_point(0).no, 0)
         If input_point_type% <> new_point_on_line Or _
             t_ele1.no <> temp_line(1) Then
         If input_point_type% <> exist_point Then  '不是旧的自由点
          Call remove_point(temp_point(0).no, display, 0)
          GoTo draw_picture1_mark2
          Else 'End If
          GoTo draw_picture1_mark2
          End If
         End If
      Call draw_tangent_line(0)
      Call draw_tangent_line(1)
      Call remove_point(temp_point(6).no, display, 0)
      temp_line(1) = line_number(temp_point(1).no, temp_point(0).no, _
                                 pointapi0, pointapi0, _
                                 depend_condition(point_, temp_point(0).no), _
                                 depend_condition(point_, temp_point(1).no), _
                                 condition, condition_color, 1, 0)
       Call set_point_name(temp_point(0).no, C_display_wenti.m_condition(num, j%))
       m_poi(temp_point(0).no).data(0).degree = 1
       Call set_point_in_line(temp_point(0).no, temp_line(1))
       Call C_display_wenti.set_m_point_no(num, _
             point_number(C_display_wenti.m_condition(num, j%)), j%, True)
     If C_display_wenti.m_no(num) = 3 Then
          If Abs(m_poi(temp_point(2).no).data(0).data0.coordinate.X - _
            m_poi(temp_point(3).no).data(0).data0.coordinate.X) > 4 Then
           A! = -(m_poi(temp_point(0).no).data(0).data0.coordinate.Y - _
             m_poi(temp_point(1).no).data(0).data0.coordinate.Y) / _
               (m_poi(temp_point(3).no).data(0).data0.coordinate.X - _
                m_poi(temp_point(2).no).data(0).data0.coordinate.X)
          Else
           A! = (m_poi(temp_point(0).no).data(0).data0.coordinate.X - _
              m_poi(temp_point(1).no).data(0).data0.coordinate.X) / _
               (m_poi(temp_point(3).no).data(0).data0.coordinate.Y - _
                m_poi(temp_point(2).no).data(0).data0.coordinate.Y)
          End If
       Call vertical_line(temp_line(0), temp_line(1), True, True) '
  ElseIf C_display_wenti.m_no(num) = 2 Then
          If Abs(m_poi(temp_point(2).no).data(0).data0.coordinate.X - _
            m_poi(temp_point(3).no).data(0).data0.coordinate.X) > 4 Then
           A! = (m_poi(temp_point(0).no).data(0).data0.coordinate.X - _
               m_poi(temp_point(1).no).data(0).data0.coordinate.X) / _
                (m_poi(temp_point(3).no).data(0).data0.coordinate.X - _
                 m_poi(temp_point(2).no).data(0).data0.coordinate.X)
          Else
           A! = (m_poi(temp_point(0).no).data(0).data0.coordinate.Y - _
               m_poi(temp_point(1).no).data(0).data0.coordinate.Y) / _
                 (m_poi(temp_point(3).no).data(0).data0.coordinate.Y - _
                    m_poi(temp_point(2).no).data(0).data0.coordinate.Y)
          End If
     End If
   
   ' Call put_name(C_display_wenti.m_point_no(num,3))
   Call C_display_wenti.set_m_point_no(num, Int(A! * 1000), 7, False)
   Call C_display_wenti.set_m_point_no(num, j% + 1, 10, False)
   'Call C_display_wenti.set_m_point_no(temp_point(2), 11)
   'Call C_display_wenti.set_m_point_no(temp_point(3), 12)
   'Call C_display_wenti.set_m_point_no(temp_point(1), 13)
   'Call C_display_wenti.set_m_point_no(temp_point(0), 14)
  Else
  temp_line(0) = line_number(C_display_wenti.m_point_no(num, 0), _
                             C_display_wenti.m_point_no(num, 1), _
                             pointapi0, pointapi0, _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                             condition, condition_color, 1, 0)
  temp_line(1) = line_number(C_display_wenti.m_point_no(num, 2), _
                             C_display_wenti.m_point_no(num, 3), _
                             pointapi0, pointapi0, _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                             condition, condition_color, 1, 0)
  If C_display_wenti.m_no(num) = 2 Then
  Call paral_line(temp_line(0), temp_line(1), True, True)
  Else
  Call vertical_line(temp_line(0), temp_line(1), True, True)
  End If
  End If
End Sub

Public Sub draw_picture_4(ByVal num As Integer, ByVal no_reduce)
'-4 与⊙□[down\\(_)]相切于点□的切线交⊙□[down\\(_)]
Dim i%
Dim ang(1) As Integer
For i% = 0 To 5
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
      C_display_wenti.m_condition(num, i%)) Then
      Exit Sub
End If
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
Next i%
Call line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                 condition, condition_color, 1, 0)
Call line_number(C_display_wenti.m_point_no(num, 2), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                 condition, condition_color, 1, 0)
Call line_number(C_display_wenti.m_point_no(num, 3), _
                 C_display_wenti.m_point_no(num, 4), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 4)), _
                 condition, condition_color, 1, 0)
Call line_number(C_display_wenti.m_point_no(num, 5), _
                 C_display_wenti.m_point_no(num, 4), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 5)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 4)), _
                 condition, condition_color, 1, 0)
End Sub

Public Sub draw_picture_5(ByVal num As Integer, ByVal no_reduce As Byte)
'-5 ∠□□□=∠□□□
'-5 ∠□□□=!_~°
Dim i%
Dim value1 As String
Dim ang As Integer
For i% = 0 To 2
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
     C_display_wenti.m_condition(num, i%)) Then
   Exit Sub
End If
Next i%
Call line_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                 condition, condition_color, 1, 0)
Call line_number(C_display_wenti.m_point_no(num, 2), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                 condition, condition_color, 1, 0)
End Sub
Public Sub draw_picture4(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%
Dim t_ele1 As condition_type
Dim t_ele2 As condition_type
Dim temp_x&, temp_y&
Dim A!
Dim temp_record As total_record_type
temp_line(0) = line_number(C_display_wenti.m_point_no(num, 0), _
                           C_display_wenti.m_point_no(num, 1), _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                           condition, condition_color, 1, 0)
For i% = 0 To 2
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
     C_display_wenti.m_condition(num, i%)) Then
      Exit Sub
End If
event_statue = 0
Exit Sub
Next i%
'If C_display_wenti.m_point_no(num,3) = 0 Then
'Call set_Divide_Point(C_display_wenti.m_point_no(num,0), _
        C_display_wenti.m_point_no(num,1), 1, 1, _
           C_display_wenti.m_point_no(num,3), True)
           'Call set_point_name(C_display_wenti.m_point_no(num,3), empty_char)
'End If
Exit Sub
 '中点
 ' 画线
 '     last_set_point = last_set_point + 1
 '      temp_point(2) = 100 - last_set_point
 '      Call set_set_point(temp_point(6))
If C_display_wenti.m_point_no(num, 2) = 0 Then
    t_coord.X = m_poi(C_display_wenti.m_point_no(num, 3)).data(0).data0.coordinate.X + _
      2 * (m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y - _
          m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y) '
    t_coord.Y = m_poi(C_display_wenti.m_point_no(num, 3)).data(0).data0.coordinate.Y - _
      2 * (m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X - _
           m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X) '
            temp_point(6).no = 0
     Call set_aid_point(temp_point(2).no, t_coord, 1)
     Call C_display_wenti.set_m_point_no(num, _
      line_number0(C_display_wenti.m_point_no(num, 3), _
           temp_point(2).no, 0, 0), 5, False)
     'lin(C_display_wenti.m_point_no(num,5)).data(0).data0.visible = 4
draw_picture1_mark4:
 event_statue = wait_for_draw_point

     While event_statue = wait_for_draw_point
     DoEvents
      Wend
 If event_statue = draw_point_down Or _
       event_statue = draw_point_move Or _
           event_statue = draw_point_up Then
   temp_x& = input_coord.X
    temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark4
      End If
     input_point_type% = read_inter_point(t_coord, t_ele1, _
                              t_ele2, temp_point(0).no, True)
     Call set_point_no_reduce(temp_point(0).no, 0)
         If input_point_type% <> new_point_on_line Or _
             t_ele1.no <> C_display_wenti.m_point_no(num, 5) Then
          If input_point_type% <> exist_point Then '不是旧的自由点
           Call remove_point(temp_point(0).no, display, 0)
          Else 'End If
           GoTo draw_picture1_mark4
          End If
         End If
 
Call remove_point(temp_point(2).no, display, 0)

  Call C_display_wenti.set_m_point_no(num, temp_point(0).no, 2, False)
  Call set_point_name(C_display_wenti.m_point_no(num, 2), _
                   C_display_wenti.m_condition(num, 2))
  m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree = 1
  Call set_point_in_line(C_display_wenti.m_point_no(num, 2), _
         C_display_wenti.m_point_no(num, 5))
   '设置新点
   temp_line(1) = line_number(temp_point(0).no, C_display_wenti.m_point_no(num, 3), _
                              pointapi0, pointapi0, _
                              depend_condition(point_, temp_point(0).no), _
                              depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                              condition, condition_color, 1, 0)
                Call vertical_line(temp_line(0), _
                 temp_line(1), True, True)
'record_0 = record0
            Call set_dverti(temp_line(0), _
                 temp_line(1), temp_record, 0, 0, True)
If Abs(m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X - _
        m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X) > 4 Then
 A! = (m_poi(C_display_wenti.m_point_no(num, 3)).data(0).data0.coordinate.Y - _
      m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.Y) / _
        (m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X - _
          m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X)
Else
 A! = (m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.X - _
    m_poi(C_display_wenti.m_point_no(num, 3)).data(0).data0.coordinate.X) / _
       (m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y - _
          m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y)
End If
  Call C_display_wenti.set_m_point_no(num, Int(A! * 1000), 7, False)

 '  Call put_name(C_display_wenti.m_point_no(num,2))
Else
 If C_display_wenti.m_point_no(num, 3) = 0 Then
  last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no = 1
   Call C_display_wenti.set_m_point_no(num, last_conditions.last_cond(1).point_no, 3, True)
    t_coord.X = _
     (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X + _
        m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X) \ 2
     Call C_display_wenti.set_m_point_no(num, last_conditions.last_cond(1).point_no, 3, False)
    t_coord.Y = _
     (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y + _
        m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y) \ 2
        Call set_point_coordinate(C_display_wenti.m_point_no(num, 3), t_coord, False)
    temp_line(1) = line_number0(C_display_wenti.m_point_no(num, 2), _
       C_display_wenti.m_point_no(num, 3), 0, 0)
   Call vertical_line(temp_line(0), temp_line(1), True, True)
   Call C_display_wenti.set_m_point_no(num, 5, 10, False)
  End If
End If
End Sub

Public Sub draw_picture5_15(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%, t_ele1%
For i% = 0 To 1
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
      C_display_wenti.m_condition(num, i%)) Then
      Exit Sub
End If
   Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
Next i%
t_ele1% = line_number(C_display_wenti.m_point_no(num, 0), _
                      C_display_wenti.m_point_no(num, 1), _
                      pointapi0, pointapi0, _
                      depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                      depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                      condition, condition_color, 1, 0)
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
MDIForm1.Toolbar1.Buttons(21).Image = 33
 '  Call init_Point0(last_conditions.last_cond(1).point_no)
 Call C_display_wenti.set_m_point_no(num, _
      last_conditions.last_cond(1).point_no, 2, False)
t_coord.X = (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X + _
   m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X) \ 2
t_coord.Y = (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y + _
   m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y) \ 2
Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
m_poi(last_conditions.last_cond(1).point_no).data(0).degree = 0
Call set_point_visible(last_conditions.last_cond(1).point_no, 1, False)
record_0.data0.condition_data.condition_no = 0
Call add_point_to_line(last_conditions.last_cond(1).point_no, t_ele1%, 0, display, _
     True, 0, temp_record)
  'poi(C_display_wenti.m_point_no(num,2)).data(0).data0.name = C_display_wenti.m_condition(2)
  'Call draw_point(Draw_form, poi(last_conditions.last_cond(1).point_no), 0, display)
If C_display_wenti.m_no(num) = 15 Then
Call add_point_to_m_circle(C_display_wenti.m_point_no(num, 0), _
       m_circle_number(1, C_display_wenti.m_point_no(num, 2), pointapi0, _
        0, 0, 0, 0, 0, 0, 1, 1, condition, condition_color, False), record0, 255)
Call add_point_to_m_circle(C_display_wenti.m_point_no(num, 1), _
       m_circle_number(1, C_display_wenti.m_point_no(num, 2), pointapi0, _
        0, 0, 0, 0, 0, 0, 1, 1, condition, condition_color, False), record0, 255)
End If
End Sub

Public Sub draw_picture6(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%
Dim value1 As Single
Dim wenti_data As wentitype
Call C_display_wenti.Get_wenti(num) '获取wenti_cond0即输入语句数据
wenti_data = wenti_cond0.data
For i% = 1 To 2
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
     C_display_wenti.m_condition(num, i%)) Then
      Exit Sub
End If
   
   Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
Next i%
value_for_draw(C_display_wenti.m_point_no(num, 3)) = _
         1 / value_string(number_string(C_display_wenti.m_point_no(num, 3))) + 1
 value1 = value_for_draw(C_display_wenti.m_point_no(num, 3))
 Call set_parent(line_, wenti_data.line_no(1), point_, wenti_data.point_no(0), paral_, wenti_data.point_no(1), wenti_data.point_no(2), _
             wenti_data.point_no(1))
If C_display_wenti.m_point_no(num, 3) <> 0 And C_display_wenti.m_point_no(num, 4) <> 0 Then
Call set_Divide_Point(C_display_wenti.m_point_no(num, 1), _
 C_display_wenti.m_point_no(num, 2), _
  value1, C_display_wenti.m_point_no(num, 0), True)
Call set_point_name(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_condition(num, 0))
'm_poi(C_display_wenti.m_point_no(num, 0)).data(0).degree = 0
'Call put_name(C_display_wenti.m_point_no(num,0))
End If
End Sub

Public Sub draw_picture7(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%
Dim t_ele2 As condition_type
Dim t_ele1 As condition_type
For i% = 0 To 1
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
     C_display_wenti.m_condition(num, i%)) Then
     Exit Sub
End If
Next i%
    Call change_point_degree(C_display_wenti.m_point_no(num, 1), -3)
    Call C_display_wenti.set_m_condition(num, "c", 7)
    Call C_display_wenti.set_m_point_no(num, _
    m_circle_number(1, C_display_wenti.m_point_no(num, 0), pointapi0, _
                  C_display_wenti.m_point_no(num, 1), 0, 0, 0, 0, 0, _
                  1, 1, condition, condition_color, True), 7, False)  '%, n%)       '两点圆
draw_picture1_mark7:
event_statue = wait_for_draw_point
While event_statue = wait_for_draw_point
DoEvents
Wend
 If event_statue = draw_point_down Or _
  event_statue = draw_point_move Or _
       event_statue = draw_point_up Then 'mouse_type <> 1 Then
   t_coord = input_coord
   ' temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark7
      End If
     input_point_type% = read_inter_point(t_coord, _
                         t_ele1, t_ele2, temp_point(0).no, True)
       Call set_point_no_reduce(temp_point(0).no, 0)
       If input_point_type% <> new_point_on_circle Or _
         t_ele1.no <> C_display_wenti.m_point_no(num, 7) Then
         If input_point_type% <> exist_point Then '不是旧的自由点
          Call remove_point(temp_point(0).no, display, 0)
         Else
          GoTo draw_picture1_mark7
         End If
       End If
 Call set_point_name(temp_point(0).no, C_display_wenti.m_condition(num, 2))
 m_poi(temp_point(0).no).data(0).degree = 1
 Call set_point_in_circle(temp_point(0).no, C_display_wenti.m_point_no(num, 7))
      Call C_display_wenti.set_m_point_no(num, temp_point(0).no, 2, False)
      Call point_position_on_circle(temp_point(0).no, _
      C_display_wenti.m_point_no(num, 7), C_display_wenti.m_point_no(num, 8), _
        C_display_wenti.m_point_no(num, 9))
 'Call put_name(temp_point(0))
End Sub

Public Sub draw_picture8(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%
For i% = 0 To 2
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
    C_display_wenti.m_condition(num, i%)) Then
    Exit Sub
End If
Next i%
   Call C_display_wenti.set_m_point_no(num, _
      m_circle_number(1, 0, pointapi0, _
              C_display_wenti.m_point_no(num, 0), _
               C_display_wenti.m_point_no(num, 1), _
                C_display_wenti.m_point_no(num, 2), _
                  0, 0, 0, 1, 1, condition, condition_color, True), 12, False)
'Call draw_circle(Draw_form, m_circ(C_display_wenti.m_point_no(num,12)).data(0).data0)
End Sub

Public Sub draw_picture9(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%
For i% = 0 To 3
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
     C_display_wenti.m_condition(num, i%)) Then
     Exit Sub
End If
   Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
Next i%
Call inter_point_line_line(C_display_wenti.m_point_no(num, 0), _
C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
 C_display_wenti.m_point_no(num, 3), 0, 0, C_display_wenti.m_point_no(num, 4), _
     pointapi0, False, True, False)
' poi(C_display_wenti.m_point_no(num,4)).data(0).data0.name = C_display_wenti.m_condition(4)
'Call put_name(C_display_wenti.m_point_no(num,4))
End Sub

Public Sub draw_picture10_16(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%
Dim t_ele1 As condition_type
Dim t_ele2 As condition_type
For i% = 0 To 4
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
      C_display_wenti.m_condition(num, i%)) Then
      Exit Sub
End If
Next i%
    temp_line(0) = _
     line_number(C_display_wenti.m_point_no(num, 1), _
                 C_display_wenti.m_point_no(num, 2), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                 condition, condition_color, 1, 0)  '(k%, l%)
    Call C_display_wenti.set_m_point_no(num, m_circle_number( _
        1, C_display_wenti.m_point_no(num, 3), pointapi0, _
         C_display_wenti.m_point_no(num, 4), 0, 0, 0, 0, 0, _
           1, 1, condition, condition_color, False), 9, False)
   Call C_display_wenti.set_m_point_no(num, _
           m_circle_number(1, C_display_wenti.m_point_no(num, 3), pointapi0, _
                      C_display_wenti.m_point_no(num, 4), 0, 0, 0, 0, 0, _
                       1, 1, condition, condition_color, True), 12, False)
    Call C_display_wenti.set_m_point_no(num, temp_line(0), 10, False)
    last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1 '***
    MDIForm1.Toolbar1.Buttons(21).Image = 33
     'Call init_Point0(last_conditions.last_cond(1).point_no)
      temp_point(0).no = last_conditions.last_cond(1).point_no
   If C_display_wenti.m_no(num) = 10 Then
   t_coord.X = m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X + _
     (m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.X - _
      m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X) * 100
   t_coord.Y = m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y + _
     (m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.Y - _
      m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y) * 100
      Call set_point_coordinate(temp_point(0).no, t_coord, False)
  temp_line(1) = line_number(C_display_wenti.m_point_no(num, 0), _
                             temp_point(0).no, pointapi0, pointapi0, _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                             depend_condition(point_, temp_point(0).no), condition, condition_color, 0, 0)
     Call paral_line(temp_line(0), temp_line(1), True, True) '
  Else
   t_coord.X = m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.X + _
     (m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.Y - _
      m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.Y) * 100
   t_coord.Y = m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate.Y - _
     (m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate.X - _
      m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate.X) * 100
      Call set_point_coordinate(temp_point(0).no, t_coord, False)
  temp_line(1) = line_number(C_display_wenti.m_point_no(num, 0), _
                             temp_point(0).no, pointapi0, pointapi0, _
                             depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                             depend_condition(point_, temp_point(0).no), _
                             condition, condition_color, 0, 0)
End If
'***********************
draw_picture1_mark10:
    event_statue = wait_for_draw_point
     While event_statue = wait_for_draw_point
     DoEvents
      Wend
 If event_statue = draw_point_down Or _
    event_statue = draw_point_move Or _
        event_statue = draw_point_up Then 'mouse_type <> 1 Then
   t_coord = input_coord
   '' temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark10
 End If
     input_point_type% = read_inter_point(t_coord, _
                             t_ele1, t_ele2, temp_point(1).no, True)
         Call set_point_no_reduce(temp_point(1).no, 0)
         If input_point_type% = new_point_on_line_circle12 And _
          t_ele1.no = temp_line(1) And t_ele2.no = C_display_wenti.m_point_no(num, 9) Then
         Call C_display_wenti.set_m_point_no(0, 1, 7, False)
         Call remove_point(temp_point(0).no, no_display, 0)
         ElseIf input_point_type% = new_point_on_line_circle21 And _
          t_ele1.no = temp_line(1) And t_ele2.no = C_display_wenti.m_point_no(num, 9) Then
          Call C_display_wenti.set_m_point_no(0, 2, 7, False)
               Call remove_point(temp_point(0).no, no_display, 0)
         Else
           If input_point_type% <> exist_point Then
            Call remove_point(temp_point(1).no, display, 0)
            End If
           GoTo draw_picture1_mark10
         End If
           Call C_display_wenti.set_m_point_no(0, temp_point(1).no, 5, True)
             Call set_line_visible(temp_line(1), 1)
         Call C_display_wenti.set_m_point_no(0, temp_point(1).no, 14, True)
         m_poi(C_display_wenti.m_point_no(num, 14)).data(0).degree = 0
    '直线与圆的一个交点
'***************************************
End Sub

Public Sub draw_picture11(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%, j%, t_point1%
Dim t_ele1 As condition_type
Dim t_ele2 As condition_type
For i% = 0 To 3
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
      C_display_wenti.m_condition(num, i%)) Then
      Exit Sub
End If
   Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
Next i%
  Call C_display_wenti.set_m_point_no(0, _
  line_number(C_display_wenti.m_point_no(num, 1), _
              C_display_wenti.m_point_no(num, 2), _
              pointapi0, pointapi0, _
              depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
              depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
              condition, condition_color, 1, 0), 10, False)
  Call C_display_wenti.set_m_point_no(0, _
    m_circle_number(1, C_display_wenti.m_point_no(num, 3), pointapi0, _
                     C_display_wenti.m_point_no(num, 4), 0, 0, 0, 0, 0, _
                      1, 1, condition, condition_color, True), 12, False)
For i% = 1 To m_lin(C_display_wenti.m_point_no(num, 10)).data(0).data0.in_point(0)
 For j% = 1 To m_Circ(C_display_wenti.m_point_no(num, 12)).data(0).data0.in_point(0)
  If m_lin(C_display_wenti.m_point_no(num, 10)).data(0).data0.in_point(i%) = _
     m_Circ(C_display_wenti.m_point_no(num, 12)).data(0).data0.in_point(j%) Then
       Call C_display_wenti.set_m_point_no(num, _
       m_Circ(C_display_wenti.m_point_no(num, 12)).data(0).data0.in_point(j%), 15, False)
    last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
    MDIForm1.Toolbar1.Buttons(21).Image = 33
     'Call init_Point0(last_conditions.last_cond(1).point_no)
   Call C_display_wenti.set_m_point_no(0, last_conditions.last_cond(1).point_no, 0, True)
   Call set_point_name(C_display_wenti.m_point_no(num, 0), _
         C_display_wenti.m_condition(num, 0))
   Call set_point_visible(C_display_wenti.m_point_no(num, 0), 1, False)
Call inter_point_line_circle2(line_number0( _
 C_display_wenti.m_point_no(num, 1), _
  C_display_wenti.m_point_no(num, 2), 0, 0), _
   C_display_wenti.m_point_no(num, 15), _
    C_display_wenti.m_point_no(num, 3), _
      t_coord, C_display_wenti.m_point_no(num, 0))
      record_0.data0.condition_data.condition_no = 0
 Call add_point_to_line(C_display_wenti.m_point_no(num, 0), _
        C_display_wenti.m_point_no(num, 10), 0, display, _
         True, 0, temp_record) 'c_display_wenti.m_point_no(num,5) = temp_point(0)
Call add_point_to_m_circle(C_display_wenti.m_point_no(num, 0), _
                      C_display_wenti.m_point_no(num, 12), record0, 255)
   ' Call draw_point(Draw_form, poi(C_display_wenti.m_point_no(num,0)), 0, display)
      Exit Sub
 End If
 Next j%
Next i%
draw_picture1_mark11:
 event_statue = wait_for_draw_point
     While event_statue = wait_for_draw_point
     DoEvents
      Wend
 If event_statue = draw_point_down Or _
     event_statue = draw_point_move Or _
        event_statue = draw_point_up Then 'mouse_type <> 1 Then
   t_coord = input_coord
   ' temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark11
End If
     Call C_display_wenti.set_m_point_no(num, C_display_wenti.m_point_no(num, 1), 0, False)
     t_point1% = C_display_wenti.m_point_no(num, 2)
      '输入线段的端点
       input_point_type% = read_inter_point(t_coord, _
          t_ele1, t_ele2, C_display_wenti.m_point_no(num, 0), True)
          Call set_point_no_reduce(C_display_wenti.m_point_no(num, 0), 0)
         If t_ele1.no = C_display_wenti.m_point_no(num, 5) And _
              t_ele2.no = C_display_wenti.m_point_no(num, 6) Then
          '非圆线交线，或非定线，或非定圆
         If input_point_type% = new_point_on_line_circle12 Then
          Call C_display_wenti.set_m_point_no(num, 1, 7, False)
         ElseIf input_point_type% = new_point_on_line_circle21 Then
          Call C_display_wenti.set_m_point_no(num, 2, 7, False)
         End If
         Call set_point_name(C_display_wenti.m_point_no(num, 0), _
                     C_display_wenti.m_condition(num, 0))
       'Call put_name(C_display_wenti.m_point_no(num,0))
         Else  '不是旧的自由点
         If input_point_type% = new_free_point Then
         Call remove_point(C_display_wenti.m_point_no(num, 0), display, 0)
         End If
          GoTo draw_picture1_mark11
         End If
If is_point_in_line3(C_display_wenti.m_point_no(num, 3), _
    m_lin(C_display_wenti.m_point_no(num, 5)).data(0).data0, 0) Then
     Call set_point_no_reduce(C_display_wenti.m_point_no(num, 3), False)
End If
End Sub

Public Sub draw_picture13(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%, j%, t_point1%
Dim ele1 As condition_type
Dim ele2 As condition_type
For i% = 1 To 4
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
      C_display_wenti.m_condition(num, i%)) Then
      Exit Sub
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
End If
   If i% <> 2 Then
    Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
   End If '点poi(c_display_wenti.m_point_no(num,i%))参加推理
Next i%
 Call C_display_wenti.set_m_condition(num, "c", 12)
 Call C_display_wenti.set_m_condition(num, "c", 13)
 Call C_display_wenti.set_m_point_no(num, _
    m_circle_number(1, C_display_wenti.m_point_no(num, 1), pointapi0, _
       C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, _
        1, 1, condition, condition_color, True), 12, False) 'j%, k%)
 Call C_display_wenti.set_m_point_no(num, _
   m_circle_number(1, C_display_wenti.m_point_no(num, 3), pointapi0, _
      C_display_wenti.m_point_no(num, 4), 0, 0, 0, 0, 0, _
       1, 1, condition, condition_color, True), 13, False) 'l%, m%)
For i% = 1 To m_Circ(C_display_wenti.m_point_no(num, 5)).data(0).data0.in_point(0)
 For j% = 1 To m_Circ(C_display_wenti.m_point_no(num, 6)).data(0).data0.in_point(0)
  If m_Circ(C_display_wenti.m_point_no(num, 5)).data(0).data0.in_point(i%) = _
     m_Circ(C_display_wenti.m_point_no(num, 6)).data(0).data0.in_point(j%) Then
      Call C_display_wenti.set_m_point_no(num, _
       m_Circ(C_display_wenti.m_point_no(num, 6)).data(0).data0.in_point(j%), 10, False)
    last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
    MDIForm1.Toolbar1.Buttons(21).Image = 33
    ' Call init_Point0(last_conditions.last_cond(1).point_no)
      Call C_display_wenti.set_m_point_no(num, last_conditions.last_cond(1).point_no, 0, False)
 Call inter_point_circle_circle2(m_Circ(C_display_wenti.m_point_no(num, 5)).data(0).data0, _
        m_Circ(C_display_wenti.m_point_no(num, 6)).data(0).data0, _
          m_Circ(C_display_wenti.m_point_no(num, 6)).data(0).data0.in_point(j%), _
           t_coord, C_display_wenti.m_point_no(num, 0))
Call add_point_to_m_circle(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_point_no(num, 5), record0, 255) 'c_display_wenti.m_point_no(num,5) = temp_point(0)
 Call add_point_to_m_circle(C_display_wenti.m_point_no(num, 0), _
                  C_display_wenti.m_point_no(num, 6), 255)
   'Call draw_point(Draw_form, poi(C_display_wenti.m_point_no(num,0)), 0, display)
      Exit Sub
  End If
 Next j%
 Next i%
draw_picture1_mark13:
 event_statue = wait_for_draw_point

     While event_statue = wait_for_draw_point
     DoEvents
      Wend
 If event_statue = draw_point_down Or _
        event_statue = draw_point_move Or _
            event_statue = draw_point_up Then
               'mouse_type <> 1 Then
   t_coord = input_coord
   ' temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark13
      End If
     input_point_type% = read_inter_point(t_coord, _
        ele1, ele2, C_display_wenti.m_point_no(num, 0), True)
        Call set_point_no_reduce(C_display_wenti.m_point_no(num, 0), 0)
         If (ele1.no = C_display_wenti.m_point_no(num, 5) And _
               ele2.no = C_display_wenti.m_point_no(num, 6)) Or _
               (ele1.no = C_display_wenti.m_point_no(num, 6) And _
                  ele2.no = C_display_wenti.m_point_no(num, 5)) Then
         If input_point_type% = new_point_on_circle_circle12 Then
          Call C_display_wenti.set_m_point_no(num, 1, 7, False)
         ElseIf input_point_type% = new_point_on_circle_circle21 Then
          Call C_display_wenti.set_m_point_no(num, 2, 7, False)
         End If
        ' poi(C_display_wenti.m_point_no(num,0)).data(0).data0.name = _
           C_display_wenti.m_condition (0)
         ' Call put_name(C_display_wenti.m_point_no(num,0))
         Else  '不是旧的自由点
         If input_point_type% = new_free_point Then
         Call remove_point(C_display_wenti.m_point_no(num, 0), display, 0)
         End If
          GoTo draw_picture1_mark13
 End If
End Sub

Public Sub draw_picture14(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%, j%
For i% = 0 To 2
If draw_free_point(C_display_wenti.m_point_no(num, i%), _
     C_display_wenti.m_condition(num, i%)) Then
      Exit Sub
End If
   Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
Next i%
i% = line_number(C_display_wenti.m_point_no(num, 1), _
                 C_display_wenti.m_point_no(num, 2), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                 condition, condition_color, 1, 0)
Call orthofoot(C_display_wenti.m_point_no(num, 0), _
   C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
     C_display_wenti.m_point_no(num, 3), 0, True)
'poi(C_display_wenti.m_point_no(num,3)).data(0).data0.name = C_display_wenti.m_condition(3)
'Call put_name(C_display_wenti.m_point_no(num,3))
   j% = line_number0(C_display_wenti.m_point_no(num, 3), _
    C_display_wenti.m_point_no(num, 0), 0, 0)
End Sub


Public Sub draw_picture0(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%, t_point1%
Dim t_ele1 As condition_type
Dim t_ele2 As condition_type
'If open_record Then
'input_last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no   '输入自由点的计数
  i% = 0
While Asc(C_display_wenti.m_condition(num, i%)) > 63 And _
          Asc(C_display_wenti.m_condition(num, i%)) < 91
                                  '输入是大字母
draw_picture1_mark0:
  event_statue = wait_for_draw_point  '输点状态
   While event_statue = wait_for_draw_point '等待事件发生
    DoEvents
   Wend
 If event_statue = draw_point_down Or event_statue = _
             draw_point_move Or event_statue = _
                    draw_point_up Then 'mouse_type <> 1 Then
   t_coord = input_coord
   ' temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
 
     GoTo draw_picture1_mark0
      End If

    input_point_type% = read_inter_point(t_coord, t_ele1, _
                                    t_ele2, temp_point(0).no, True)
          Call set_point_no_reduce(temp_point(0).no, 0)
     If input_point_type% <> new_free_point Then  '不是新的自由点
         If input_point_type% <> exist_point Then  '不是旧的自由点
          Call remove_point(temp_point(0).no, display, 0) '抹掉
         Else
          GoTo draw_picture1_mark0
         End If
     End If
      ' poi(temp_point(0)).data(0).data0.name = C_display_wenti.m_condition(i%)
         Call C_display_wenti.set_m_point_no(num, temp_point(0).no, i%, False)
       ' Call put_name(temp_point(0))
   i% = i% + 1
   Wend
End Sub
Public Sub draw_picture1(ByVal num As Integer, ByVal no_reduce As Byte)
Dim i%, t_point1%
Dim t_ele1 As condition_type
Dim t_ele2 As condition_type
 For i% = 0 To 1
  If draw_free_point(C_display_wenti.m_point_no(num, i%), _
     C_display_wenti.m_condition(num, i%)) Then
     Exit Sub
  End If
   Call change_point_degree(C_display_wenti.m_point_no(num, i%), -3)
Next i%
    Call C_display_wenti.set_m_point_no(num, _
       line_number(C_display_wenti.m_point_no(num, 0), _
                   C_display_wenti.m_point_no(num, 1), _
                   pointapi0, pointapi0, _
                   depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                   depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                   condition, condition_color, 1, 0), 4, False)  '(j%, k%) ' 确定两点所在的直线
'If open_record = True Then
   last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
   MDIForm1.Toolbar1.Buttons(21).Image = 33
   'Call init_Point0(last_conditions.last_cond(1).point_no)
    Call C_display_wenti.set_m_point_no(num, last_conditions.last_cond(1).point_no, 2, False)
    'poi(last_conditions.last_cond(1).point_no).data(0).data0.name = C_display_wenti.m_condition(2)
     t_coord.X = temp_record_poi(last_conditions.last_cond(1).point_no, 0)
     t_coord.Y = temp_record_poi(last_conditions.last_cond(1).point_no, 1)
     Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
     m_poi(last_conditions.last_cond(1).point_no).data(0).degree = 1
     Call set_point_visible(last_conditions.last_cond(1).point_no, 1, False)
       ' Call draw_point(Draw_form, poi(last_conditions.last_cond(1).point_no), 0, display)
       record_0.data0.condition_data.condition_no = 0
          Call add_point_to_line(last_conditions.last_cond(1).point_no, C_display_wenti.m_point_no(num, 4), _
           0, display, True, 0, temp_record)
'Else
'***********************
draw_picture1_mark1:
   event_statue = wait_for_draw_point
    While event_statue = wait_for_draw_point
      DoEvents
    Wend
 If event_statue = draw_point_down Or _
      event_statue = draw_point_move Or _
         event_statue = draw_point_up Then 'mouse_type <> 1 Then
   t_coord = input_coord
    'temp_y& = input_coord.Y
 ElseIf event_statue = wait_for_input_char Then
   Exit Sub
 Else
     GoTo draw_picture1_mark1
 End If
    input_point_type% = read_inter_point(t_coord, t_ele1, _
                                  t_ele2, temp_point(0).no, True)
       Call set_point_no_reduce(temp_point(0).no, 0)
      If input_point_type% <> new_point_on_line Or _
         t_ele1.no <> C_display_wenti.m_point_no(num, 4) Then
         If input_point_type% <> exist_point Then  '不是旧的自由点
          Call remove_point(temp_point(0).no, display, 0)
         GoTo draw_picture1_mark1
          Else
         GoTo draw_picture1_mark1
         End If
       End If
       Call C_display_wenti.set_m_point_no(num, temp_point(0).no, 2, True)
       Call set_point_name(C_display_wenti.m_point_no(num, 2), _
          C_display_wenti.m_condition(num, 2))
       Call set_point_in_line(C_display_wenti.m_point_no(num, 2), _
               C_display_wenti.m_point_no(num, 4))
       m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree = 1
      ' Call put_name(C_display_wenti.m_point_no(num,2))
 Call set_wenti1(num, C_display_wenti.m_point_no(num, 0), _
  C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
   C_display_wenti.m_point_no(num, 4))
End Sub

Public Sub draw_polygon_sides_for_inform(n() As Integer, no As Byte, co As Long)
Dim i%
If line_width < 2 Then
Draw_form.DrawWidth = 2
End If
For i% = 0 To no - 2
 Call Drawline(Draw_form, co, 0, _
       m_poi(n(i%)).data(0).data0.coordinate, m_poi(n(i% + 1)).data(0).data0.coordinate, 0)
Next i%
 Call Drawline(Draw_form, co, 0, _
       m_poi(n(0)).data(0).data0.coordinate, m_poi(n(no - 1)).data(0).data0.coordinate, 0)
       Draw_form.DrawWidth = line_width
End Sub
Public Sub draw_epolygon_for_inform(poly As epolygon_data_type, co As Long)
 Call draw_polygon_sides_for_inform(poly.p.v(), poly.p.total_v, co)
'*** Call fill_color_for_polygon(poly.p.v(), poly.p.total_v, co, co + 4)
End Sub
Public Sub draw_line_for_inform(p1%, p2%, draw_ty As Byte)
Dim l%
If draw_ty = 1 Then
Call line_number(p1%, p2%, pointapi0, pointapi0, _
                  depend_condition(0, 0), depend_condition(0, 0), _
                   conclusion, 13, 1, 0, 1)
Else
 l% = line_number0(p1%, p2%, 0, 0)
  Call C_display_picture.re_draw_line(l%)
End If
End Sub


Public Sub draw_picture22(ByVal num As Integer)
Dim n(5) As Integer
Dim i%, k%
Dim tn(5) As Integer
Dim ts(4) As String
Dim tp(3) As Integer
Dim temp_record As total_record_type
Dim it(2) As Integer
Dim ite(2) As item0_data_type
temp_record.record_.display_no = -(num + 1)
temp_record.record_data.data0.condition_data.condition_no = 0
temp_record.record_data.data0.theorem_no = -2
k% = 0
n(0) = 0
For i% = 0 To 50
 If Asc(C_display_wenti.m_condition(num, i%)) = 13 Then
   k% = k% + 1
   tn(k%) = i%
    If k% = 5 Then
     GoTo draw_picture22_out1:
    End If
 ElseIf C_display_wenti.m_condition(num, i%) = empty_char Then
    GoTo draw_picture22_out1:
 Else
  ts(k%) = ts(k%) & C_display_wenti.m_condition(num, i%)
 End If
Next i%
draw_picture22_out1:
If ts(2) = "" Then
   ts(2) = "1"
ElseIf ts(2) = "-" Then
   ts(2) = "-1"
End If
If ts(3) = "" Then
   ts(3) = "1"
ElseIf ts(3) = "-" Then
   ts(3) = "-1"
End If
If ts(4) = "" Then
   ts(4) = "1"
ElseIf ts(2) = "-" Then
   ts(4) = "-1"
End If
Call read_element_from_wenti(num, tn(0), tn(1) - 1, tp(0), tp(1), 0)
Call read_element_from_wenti(num, tn(1) + 1, tn(2) - 1, tp(2), tp(3), 0)
ts(2) = initial_string(ts(2))
ts(3) = initial_string(ts(3))
ts(4) = initial_string(ts(4))
If solut_2order_equation(ts(2), ts(3), ts(4), ts(0), ts(1), False) Then
   If tp(1) > 0 And tp(3) > 0 Then
          If squre_distance_point_point(m_poi(tp(0)).data(0).data0.coordinate, _
           m_poi(tp(1)).data(0).data0.coordinate) > _
             squre_distance_point_point(m_poi(tp(2)).data(0).data0.coordinate, _
                m_poi(tp(3)).data(0).data0.coordinate) Then
          Call set_line_value(tp(0), tp(1), ts(0), 0, 0, 0, temp_record.record_data, 0, 0, False)
          Call set_line_value(tp(2), tp(3), ts(1), 0, 0, 0, temp_record.record_data, 0, 0, False)
          Else
          Call set_line_value(tp(0), tp(1), ts(1), 0, 0, 0, temp_record.record_data, 0, 0, False)
          Call set_line_value(tp(2), tp(3), ts(0), 0, 0, 0, temp_record.record_data, 0, 0, False)
          End If
   Else
   End If
Else
   If tp(1) > 0 And tp(3) > 0 Then
    Call set_two_line_value(tp(0), tp(1), tp(2), tp(3), _
        0, 0, 0, 0, 0, 0, "1", "1", divide_string(time_string("-1", ts(3), False, False), ts(2), True, False), _
         temp_record, 0, 0)
   Else
   End If
    Call set_item0(tp(0), tp(1), tp(2), tp(3), "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", _
          divide_string(ts(4), ts(2), True, False), 0, temp_record.record_data.data0.condition_data, _
              0, it(0), 0, 0, condition_data0, False)
    Call set_general_string(it(0), it(1), 0, 0, "1", "1", "0", "0", _
      ts(1), 0, 0, 0, temp_record, 0, 0)
End If
End Sub
Public Sub read_element_from_wenti(w_n%, s%, l%, p1%, p2%, order%)
Dim nt%
If l% > 0 And s% > l% Then
 Exit Sub
Else
 If C_display_wenti.m_condition(w_n%, s%) = "(" And _
      C_display_wenti.m_condition(w_n%, l%) = ")" Then
   Call read_element_from_wenti(w_n%, s% + 1, l% - 1, p1%, p2%, order%)
 ElseIf C_display_wenti.m_condition(w_n%, s%) = "(" Then
   If l% - s% > 2 Then
    If C_display_wenti.m_condition(w_n%, l% - 2) = ")" Then
     If C_display_wenti.m_condition(w_n%, l% - 1) = "^" Then
       order% = val(C_display_wenti.m_condition(w_n%, l%))
       Call read_element_from_wenti(w_n%, s% + 1, l% - 1, p1%, p2%, 0)
     End If
   End If
  End If
 Else
  If C_display_wenti.m_point_no(w_n%, s%) > 0 And C_display_wenti.m_point_no(w_n%, s% + 1) > 0 Then
   p1% = C_display_wenti.m_point_no(w_n%, s%)
   p2% = C_display_wenti.m_point_no(w_n%, s% + 1)
   nt% = s% + 2
  Else
    If C_display_wenti.m_condition(w_n%, s%) = "$" Then 'sin
     p2% = -1
    ElseIf C_display_wenti.m_condition(w_n%, s%) = "&" Then 'cos
     p2% = -2
    ElseIf C_display_wenti.m_condition(w_n%, s%) = "`" Then 'tan
     p2% = -3
    ElseIf C_display_wenti.m_condition(w_n%, s%) = "\" Then 'ctan
     p2% = -4
    ElseIf C_display_wenti.m_condition(w_n%, s%) = Chr(12) Then '"," Then
     p2% = -6
    ElseIf C_display_wenti.m_condition(w_n%, s%) = "?" Then '向量
     p2% = -10
    End If
    If p2% = -6 Then
     p1% = Abs(angle_number(C_display_wenti.m_point_no(w_n%, s% + 1), _
         C_display_wenti.m_point_no(w_n%, s% + 2), C_display_wenti.m_point_no(w_n%, s% + 3), 0, 0))
     nt% = s% + 4
    ElseIf p2% = -10 And C_display_wenti.m_condition(w_n%, s% + 3) = "}" Then
     p1% = vector_number(C_display_wenti.m_point_no(w_n%, s% + 1), C_display_wenti.m_point_no(w_n%, s% + 2), 0)
     nt% = s% + 3
    Else
     p1% = Abs(angle_number(C_display_wenti.m_point_no(w_n%, s% + 1), _
         C_display_wenti.m_point_no(w_n%, s% + 2), C_display_wenti.m_point_no(w_n%, s% + 3), 0, 0))
    nt% = s% + 4
    End If
 End If
 If nt% < l% Then
  If C_display_wenti.m_condition(w_n%, nt%) = "^" Then
       order% = val(C_display_wenti.m_condition(w_n%, nt% + 1))
  End If
 Else
  order% = 1
 End If
 End If
 End If
End Sub

Public Sub con_relation_(conc_no%, rela As relation_data0_type)
If conclusion_data(conc_no%).ty = midpoint_ Then
 con_mid_point(conc_no%).data(0).line_no = con_relation(conc_no%).data(0).line_no(0)
 con_mid_point(conc_no%).data(0).n(0) = con_relation(conc_no%).data(0).n(0)
 con_mid_point(conc_no%).data(0).n(1) = con_relation(conc_no%).data(0).n(1)
 con_mid_point(conc_no%).data(0).n(2) = con_relation(conc_no%).data(0).n(3)
 con_mid_point(conc_no%).data(0).poi(0) = con_relation(conc_no%).data(0).poi(0)
 con_mid_point(conc_no%).data(0).poi(1) = con_relation(conc_no%).data(0).poi(1)
 con_mid_point(conc_no%).data(0).poi(2) = con_relation(conc_no%).data(0).poi(3)
ElseIf conclusion_data(conc_no%).ty = eline_ Then
 con_eline(conc_no%).data(0).data0.line_no(0) = con_relation(conc_no%).data(0).line_no(0)
 con_eline(conc_no%).data(0).data0.line_no(1) = con_relation(conc_no%).data(0).line_no(1)
 con_eline(conc_no%).data(0).data0.n(0) = con_relation(conc_no%).data(0).n(0)
 con_eline(conc_no%).data(0).data0.n(1) = con_relation(conc_no%).data(0).n(1)
 con_eline(conc_no%).data(0).data0.n(2) = con_relation(conc_no%).data(0).n(2)
 con_eline(conc_no%).data(0).data0.n(3) = con_relation(conc_no%).data(0).n(3)
 con_eline(conc_no%).data(0).data0.poi(0) = con_relation(conc_no%).data(0).poi(0)
 con_eline(conc_no%).data(0).data0.poi(1) = con_relation(conc_no%).data(0).poi(1)
 con_eline(conc_no%).data(0).data0.poi(2) = con_relation(conc_no%).data(0).poi(2)
 con_eline(conc_no%).data(0).data0.poi(3) = con_relation(conc_no%).data(0).poi(3)
End If
End Sub

Public Sub string_from_wenti_condition(w_n%, str() As String)
Dim i%, j%
i% = 0
For i% = 0 To 50
If C_display_wenti.m_condition(w_n%, i%) = ";" Then
j% = j% + 1
ElseIf Asc(C_display_wenti.m_condition(w_n%, i%)) = 13 Then
 Exit Sub
Else
str(j%) = str(j%) + C_display_wenti.m_condition(w_n%, i%)
End If
Next i%
End Sub
Public Function from_string_to_temp_item(s As String, para As String) As Integer
Dim temp_s(3) As String
Dim i%, j%, n%
Dim brace As Integer
Dim tn(1) As Integer
Dim tp(2) As Integer
Dim ch(1) As String * 1
Dim re_condition As condition_data_type
For i% = 1 To Len(s)
 ch(0) = Mid$(s, i%, 1)
  If ch(0) = "+" And i% > 2 Then
    If Mid$(s, i% - 2, 1) >= "A" And Mid$(s, i% - 1, 1) = ")" Then
      tn(0) = from_string_to_temp_item(Mid$(s, 1, i% - 1), temp_s(0))
        tn(1) = from_string_to_temp_item(Mid$(s, i% + 1, Len(s)), temp_s(1))
           Call set_temp_item0(tn(0), -7, tn(1), -7, "+", _
             temp_s(0), temp_s(1), para$, from_string_to_temp_item)
               Exit Function
    End If
  ElseIf (ch(0) = "-1" Or ch(0) = "@1") And i% > 2 Then
    If Mid$(s, i% - 2, 1) >= "A" And Mid$(s, i% - 1, 1) = ")" Then
     brace = 0
      temp_s(2) = ""
      For j% = i% + 1 To Len(s)
       ch(1) = Mid$(s, j%, 1)
        If ch(1) = "(" Then
          brace = brace + 1
        ElseIf ch(1) = ")" Then
          brace = brace - 1
        ElseIf ch(1) = "-" And brace = 0 Then
          temp_s(2) = temp_s(2) + "+"
        ElseIf ch(1) = "+" And brace = 0 Then
         temp_s(2) = temp_s(2) + "-"
        Else
          temp_s(2) = temp_s(2) + ch(1)
        End If
     Next j%
       tn(0) = from_string_to_temp_item(Mid$(s, 1, i% - 1), temp_s(0))
       tn(1) = from_string_to_temp_item(temp_s(2), temp_s(1))
       Call set_temp_item0(tn(0), -7, tn(1), -7, "-", _
           temp_s(0), temp_s(1), para$, from_string_to_temp_item)
             Exit Function
     End If
  ElseIf ch(0) = "*" And i% > 2 Then
   If Mid$(s, i% - 2, 1) > "A" And Mid$(s, i% - 1, 1) > ")" Then
     tn(0) = from_string_to_temp_item(Mid$(s, 1, i% - 1), temp_s(0))
     tn(1) = from_string_to_temp_item(Mid$(s, i% + 1, Len(s)), temp_s(1))
      Call set_temp_item0(tn(0), -7, tn(1), -7, "*", _
           temp_s(0), temp_s(1), para$, from_string_to_temp_item)
             Exit Function
   End If
 ElseIf ch(0) = "/" And i% > 2 Then
   If Mid$(s, i% - 2, 1) > "A" And Mid$(s, i% - 1, 1) > ")" Then
    tn(0) = from_string_to_temp_item(Mid$(s, 1, i% - 1), temp_s(0))
    tn(1) = from_string_to_temp_item(Mid$(s, i% + 1, Len(s)), temp_s(1))
    Call set_temp_item0(tn(0), -7, tn(1), -7, "/", _
           temp_s(0), temp_s(1), para$, from_string_to_temp_item)
             Exit Function
   End If
 End If
Next i%
'*************************************************************
temp_s(2) = ""
If Mid$(s, Len(s), 1) = ")" Then
   brace = 1
   For i% = Len(s) - 1 To 1 Step -1
   ch(0) = Mid$(s, i%, 1)
   If ch(0) = ")" Then
   brace = brace + 1
   ElseIf ch(0) = "(" Then
   brace = brace - 1
   ElseIf brace <> 0 Then
   temp_s(2) = ch(2) + temp_s(2)
   Else
   temp_s(3) = Mid$(s, 1, i% - 1)
   GoTo from_string_to_temp_item_mark0
   End If
   Next i%
End If
from_string_to_temp_item_mark0:
s = temp_s(2)
If temp_s(3) = "" Then
temp_s(3) = "1"
End If
For i% = 1 To Len(s)
 ch(0) = Mid$(s, i%, 1)
  If ch(0) = "+" And i% > 1 Then
    If Mid$(s, i% - 1, 1) >= "A" Then
      tn(0) = from_string_to_temp_item(Mid$(s, 1, i% - 1), temp_s(0))
        tn(1) = from_string_to_temp_item(Mid$(s, i% + 1, Len(s)), temp_s(1))
           Call set_temp_item0(tn(0), -7, tn(1), -7, "+", _
             temp_s(0), temp_s(1), para$, from_string_to_temp_item)
              para$ = time_string(para$, temp_s(3), True, False)
               Exit Function
    End If
  ElseIf ch(0) = "-" And i% > 1 Then
    If Mid$(s, i% - 1, 1) >= "A" Then
     temp_s(2) = ""
      For j% = i% + 1 To Len(s)
       ch(1) = Mid$(s, j%, 1)
        If ch(1) = "-" And Mid$(s, j% - 1, 1) = ")" Then
          temp_s(2) = temp_s(2) + "+"
        ElseIf ch(1) = "+" And Mid$(s, j% - 1, 1) = ")" Then
         temp_s(2) = temp_s(2) + "-"
        Else
          temp_s(2) = temp_s(2) + ch(1)
        End If
     Next j%
       tn(0) = from_string_to_temp_item(Mid$(s, 1, i% - 1), temp_s(0))
       tn(1) = from_string_to_temp_item(temp_s(2), temp_s(1))
       Call set_temp_item0(tn(0), -7, tn(1), -7, "-", _
           temp_s(0), temp_s(1), para$, from_string_to_temp_item)
              para$ = time_string(para$, temp_s(3), True, False)
             Exit Function
     End If
  ElseIf ch(0) = "*" And i% > 1 Then
   If Mid$(s, i% - 1, 1) > "A" Then
     tn(0) = from_string_to_temp_item(Mid$(s, 1, i% - 1), temp_s(0))
     tn(1) = from_string_to_temp_item(Mid$(s, i% + 1, Len(s)), temp_s(1))
      Call set_temp_item0(tn(0), -7, tn(1), -7, "*", _
           temp_s(0), temp_s(1), para$, from_string_to_temp_item)
              para$ = time_string(para$, temp_s(3), True, False)
             Exit Function
   End If
 ElseIf ch(0) = "/" And i% > 1 Then
   If Mid$(s, i% - 1, 1) > "A" Then
    tn(0) = from_string_to_temp_item(Mid$(s, 1, i% - 1), temp_s(0))
    tn(1) = from_string_to_temp_item(Mid$(s, i% + 1, Len(s)), temp_s(1))
    Call set_temp_item0(tn(0), -7, tn(1), -7, "/", _
           temp_s(0), temp_s(1), para$, from_string_to_temp_item)
               para$ = time_string(para$, temp_s(3), True, False)
            Exit Function
   End If
 End If
Next i%
'*************************************************************
 
 n% = InStr(1, s, LoadResString_(1380, ""), 0)
 If n% > 0 Then
  tp(0) = point_number(Mid$(s, n% + 1, 1))
  tp(1) = point_number(Mid$(s, n% + 2, 1))
  tp(2) = point_number(Mid$(s, n% + 3, 1))
  tp(0) = Abs(angle_number(tp(0), tp(1), tp(2), 0, 0))
    n% = InStr(1, s, "sin", 0)
     If n% > 0 Then
          Call set_temp_item0(tp(0), -1, 0, 0, "~", _
           "1", "0", "", from_string_to_temp_item)
             If n% > 1 Then
               para$ = Mid$(s, 1, n% - 1)
                If para$ = "-" Then
                 para$ = "-1"
                ElseIf para$ = "+" Or para$ = "+1" Then
                 para$ = "1"
                End If
             Else
               para$ = "1"
             End If
            Exit Function
     Else
       n% = InStr(1, s, "cos", 0)
       If n% > 0 Then
          Call set_temp_item0(tp(0), -2, 0, 0, "~", _
           "1", "0", "", from_string_to_temp_item)
             If n% > 1 Then
               para$ = Mid$(s, 1, n% - 1)
                If para$ = "-" Then
                 para$ = "-1"
                ElseIf para$ = "+" Or para$ = "+1" Then
                 para$ = "1"
                End If
             Else
               para$ = "1"
             End If
            Exit Function
       Else
        n% = InStr(1, s, "tan", 0)
        If n% > 0 Then
                  Call set_temp_item0(tp(0), -3, 0, 0, "~", _
           "1", "0", "", from_string_to_temp_item)
             If n% > 1 Then
               para$ = Mid$(s, 1, n% - 1)
                If para$ = "-" Then
                 para$ = "-1"
                ElseIf para$ = "+" Or para$ = "+1" Then
                 para$ = "1"
                End If
             Else
               para$ = "1"
             End If
            Exit Function
        Else
         n% = InStr(1, s, "ctan", 0)
          If n% > 0 Then
            Call set_temp_item0(tp(0), -3, 0, 0, "~", _
           "1", "0", "", from_string_to_temp_item)
             If n% > 1 Then
               para$ = Mid$(s, 1, n% - 1)
                If para$ = "-" Then
                 para$ = "-1"
                ElseIf para$ = "+" Or para$ = "+1" Then
                 para$ = "1"
                End If
             Else
               para$ = "1"
             End If
            Exit Function

          End If
        End If
       End If
     End If
 Else
   n% = InStr(1, s, LoadResString_(1385, ""), 0)
   If n% > 0 Then
   Else
    ch(0) = Mid$(s, Len(s) - 1, 1)
    ch(1) = Mid$(s, Len(s), 1)
    tp(0) = point_number(ch(0))
    tp(1) = point_number(ch(1))
    Call set_temp_item0(tp(0), tp(1), 0, 0, "~", _
           "1", "0", para$, from_string_to_temp_item)
    If Len(s) > 2 Then
     ch(2) = Mid$(s, 1, Len(s) - 2)
    Else
     ch(2) = "1"
    End If
    If ch(2) = "-" Or ch(2) = "@" Then
     ch(2) = "-1"
    ElseIf ch(2) = "+" Or ch(2) = "+1" Then
     ch(2) = "1"
    End If
    para$ = ch(2)
   End If
End If
End Function
Public Function simple_temp_item(ByVal n%) As Integer
Dim tn(1) As Integer
Dim tp(3) As Integer
Dim re_condition As condition_data_type
Dim temp_ele1() As element_data_type
Dim temp_ele2() As element_data_type
Dim last_temp_ele1%, last_temp_ele2%
Dim i%, j%, k%
If read_element_from_temp_item(n%, temp_ele1(), last_temp_ele1%, _
          temp_ele2(), last_temp_ele2%) = 1 Then
 While i% < last_temp_ele1%
  While j% < last_temp_ele2%
   If temp_ele1(i%).poi(0) = temp_ele2(j%).poi(0) And _
        temp_ele1(i%).poi(1) = temp_ele2(j%).poi(1) Then
         last_temp_ele1% = last_temp_ele1% - 1
          For k% = i% To last_temp_ele1%
           temp_ele1(k%) = temp_ele1(k% + 1)
          Next k%
         last_temp_ele2% = last_temp_ele2% - 1
          For k% = i% To last_temp_ele2%
           temp_ele2(k%) = temp_ele2(k% + 1)
          Next k%
         i% = i% - 1
    Else
     j% = j% + 1
    End If
  Wend
   i% = i% + 1
 Wend
 If last_temp_ele1% = 0 And last_temp_ele2% = 0 Then
  simple_temp_item = 0
 ElseIf last_temp_ele1% <= 1 And last_temp_ele2% <= 1 Then
  If last_temp_ele1% = 1 And last_temp_ele2% = 1 Then
      Call set_item0(temp_ele1(0).poi(0), temp_ele1(0).poi(1), _
            temp_ele2(0).poi(0), temp_ele2(0).poi(1), _
             "/", 0, 0, 0, 0, 0, 0, temp_item0(n%).para(0), temp_item0(n%).para(1), _
              "1", "", "1", 0, re_condition, _
               0, simple_temp_item, 0, 0, condition_data0, False)
  ElseIf last_temp_ele1% = 1 Then
      Call set_item0(temp_ele1(0).poi(0), temp_ele1(0).poi(1), _
            0, 0, "~", 0, 0, 0, 0, 0, 0, temp_item0(n%).para(0), temp_item0(n%).para(1), _
              "1", "", "1", 0, re_condition, _
               0, simple_temp_item, 0, 0, condition_data0, False)
  ElseIf last_temp_ele2% = 1 Then
      Call set_item0(0, 0, temp_ele2(0).poi(0), temp_ele2(0).poi(1), _
             "/", 0, 0, 0, 0, 0, 0, temp_item0(n%).para(0), temp_item0(n%).para(1), _
              "1", "", "1", 0, re_condition, _
               0, simple_temp_item, 0, 0, condition_data0, False)
  End If
 Else
  If last_temp_ele1% = 0 Then
   tp(0) = 0
   tp(1) = 0
  ElseIf last_temp_ele1% = 1 Then
   tp(0) = temp_ele1(0).poi(0)
   tp(1) = temp_ele1(0).poi(1)
  Else
   tp(0) = from_element_to_item0(temp_ele1(), last_temp_ele1%)
   If tp(0) > 0 Then
      tp(1) = -7
   Else
      tp(1) = 0
   End If
  End If
  If last_temp_ele2% = 0 Then
   tp(2) = 0
   tp(3) = 0
  ElseIf last_temp_ele2% = 1 Then
   tp(2) = temp_ele2(0).poi(0)
   tp(3) = temp_ele2(0).poi(1)
  Else
   tp(2) = from_element_to_item0(temp_ele2(), last_temp_ele2%)
   If tp(2) > 0 Then
      tp(3) = -7
   Else
      tp(3) = 0
   End If
  End If
  If tp(0) > 0 And tp(2) > 0 Then
        Call set_item0(tp(0), tp(1), tp(2), tp(3), temp_item0(n%).sig, 0, 0, 0, 0, 0, 0, _
          temp_item0(n%).para(0), temp_item0(n%).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_temp_item, 0, 0, condition_data0, False)
  ElseIf tp(0) > 0 Then
        Call set_item0(tp(0), tp(1), 0, 0, "~", 0, 0, 0, 0, 0, 0, _
          temp_item0(n%).para(0), temp_item0(n%).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_temp_item, 0, 0, condition_data0, False)
  ElseIf tp(2) > 0 Then
   If temp_item0(n%).sig = "*" Then
        Call set_item0(tp(2), tp(3), 0, 0, "~", 0, 0, 0, 0, 0, 0, _
          temp_item0(n%).para(0), temp_item0(n%).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_temp_item, 0, 0, condition_data0, False)
   Else
        Call set_item0(0, 0, tp(2), tp(3), temp_item0(n%).sig, 0, 0, 0, 0, 0, 0, _
          temp_item0(n%).para(0), temp_item0(n%).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_temp_item, 0, 0, condition_data0, False)
   End If
  Else
   simple_temp_item = 0
  End If
 End If
Else
If temp_item0(n%).poi(1) = -7 And temp_item0(n%).poi(3) = -7 Then
   If temp_item0(temp_item0(n%).poi(0)).sig = "~" And temp_item0(temp_item0(n%).poi(2)).sig = "~" Then
    Call set_item0(temp_item0(temp_item0(n%).poi(0)).poi(0), temp_item0(temp_item0(n%).poi(0)).poi(1), _
            temp_item0(temp_item0(n%).poi(2)).poi(0), temp_item0(temp_item0(n%).poi(2)).poi(1), _
             temp_item0(n%).sig, 0, 0, 0, 0, 0, 0, temp_item0(n%).para(0), temp_item0(n%).para(1), _
              "1", "", "1", 0, re_condition, _
               0, simple_temp_item, 0, 0, condition_data0, False)
   ElseIf temp_item0(temp_item0(n%).poi(0)).sig = "~" And temp_item0(n%).sig <> "~" Then
    tn(0) = simple_temp_item(temp_item0(n%).poi(2))
     If item0(n%).data(0).sig = "~" Then
      Call set_item0(temp_item0(temp_item0(n%).poi(0)).poi(0), temp_item0(temp_item0(n%).poi(0)).poi(1), _
            item0(tn(0)).data(0).poi(0), item0(tn(0)).data(0).poi(1), _
             temp_item0(n%).sig, 0, 0, 0, 0, 0, 0, temp_item0(n%).para(0), temp_item0(n%).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_temp_item, 0, 0, condition_data0, False)
     Else
      Call set_item0(temp_item0(temp_item0(n%).poi(0)).poi(0), temp_item0(temp_item0(n%).poi(0)).poi(1), _
            tn(0), -7, temp_item0(n%).sig, 0, 0, 0, 0, 0, 0, temp_item0(n%).para(0), temp_item0(n%).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_temp_item, 0, 0, condition_data0, False)
     End If
   ElseIf temp_item0(temp_item0(n%).poi(2)).sig = "~" And temp_item0(n%).sig <> "~" Then
    tn(0) = simple_temp_item(temp_item0(n%).poi(0))
     If item0(n%).data(0).sig = "~" Then
      Call set_item0(item0(tn(0)).data(0).poi(0), item0(tn(0)).data(0).poi(1), _
          temp_item0(temp_item0(n%).poi(2)).poi(0), temp_item0(temp_item0(n%).poi(2)).poi(1), _
         temp_item0(n%).sig, 0, 0, 0, 0, 0, 0, temp_item0(n%).para(0), temp_item0(n%).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_temp_item, 0, 0, condition_data0, False)
     Else
      Call set_item0(tn(0), -7, _
          temp_item0(temp_item0(n%).poi(2)).poi(0), temp_item0(temp_item0(n%).poi(2)).poi(1), _
         temp_item0(n%).sig, 0, 0, 0, 0, 0, 0, temp_item0(n%).para(0), temp_item0(n%).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_temp_item, 0, 0, condition_data0, False)
     End If
   End If
ElseIf temp_item0(n%).poi(1) = -7 Then
 tn(0) = simple_temp_item(temp_item0(n%).poi(0))
  If item0(tn(0)).data(0).sig = "~" Then
  tp(0) = item0(tn(0)).data(0).poi(0)
  tp(1) = item0(tn(0)).data(0).poi(1)
  Else
  tp(0) = tn(0)
  tp(1) = -7
  End If
 If temp_item0(temp_item0(n%).poi(0)).sig = "~" Then
  
 Else
 End If
ElseIf temp_item0(n%).poi(3) = -7 Then
 tn(0) = simple_temp_item(temp_item0(n%).poi(2))
  If item0(tn(0)).data(0).sig = "~" Then
  tp(0) = item0(tn(0)).data(0).poi(0)
  tp(1) = item0(tn(0)).data(0).poi(1)
  Else
  tp(0) = tn(0)
  tp(1) = -7
  End If
 If temp_item0(temp_item0(n%).poi(2)).sig = "~" Then
      Call set_item0(item0(tn(0)).data(0).poi(0), item0(tn(0)).data(0).poi(1), _
          temp_item0(temp_item0(n%).poi(2)).poi(0), temp_item0(temp_item0(n%).poi(2)).poi(1), _
         temp_item0(n%).sig, 0, 0, 0, 0, 0, 0, temp_item0(n%).para(0), temp_item0(n%).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_temp_item, 0, 0, condition_data0, False)
 Else
 End If
Else
 tn(0) = simple_temp_item(temp_item0(n%).poi(0))
   If item0(tn(0)).data(0).sig = "~" Then
    tp(0) = item0(tn(0)).data(0).poi(0)
    tp(1) = item0(tn(0)).data(0).poi(1)
   Else
    tp(0) = tn(0)
    tp(1) = -7
   End If
  tn(1) = simple_temp_item(temp_item0(n%).poi(1))
   If item0(tn(0)).data(0).sig = "~" Then
    tp(2) = item0(tn(1)).data(0).poi(0)
    tp(3) = item0(tn(1)).data(0).poi(1)
   Else
    tp(0) = tn(1)
    tp(1) = -7
   End If
 End If
End If
End Function
Public Function read_element_from_temp_item(ByVal n%, ele1() As element_data_type, last_ele1%, _
              ele2() As element_data_type, last_ele2%) As Byte
Dim temp_ele1() As element_data_type
Dim last_temp_ele1%
Dim temp_ele2() As element_data_type
Dim last_temp_ele2%
Dim i%
If temp_item0(n%).sig = "~" Then
ReDim Preserve ele1(last_ele1%) As element_data_type
ele1(last_ele1%).poi(0) = temp_item0(n%).poi(0)
ele1(last_ele1%).poi(1) = temp_item0(n%).poi(1)
last_ele1% = last_ele1% + 1
read_element_from_temp_item = 1
ElseIf temp_item0(n%).sig = "*" Then
 If temp_item0(n%).poi(1) <> -7 And temp_item0(n%).poi(3) <> -7 Then
  ReDim Preserve ele1(last_ele1%) As element_data_type
  ele1(last_ele1%).poi(0) = temp_item0(n%).poi(0)
  ele1(last_ele1%).poi(1) = temp_item0(n%).poi(1)
  last_ele1% = last_ele1% + 1
  ReDim Preserve ele1(last_ele1%) As element_data_type
  ele1(last_ele1%).poi(0) = temp_item0(n%).poi(2)
  ele1(last_ele1%).poi(1) = temp_item0(n%).poi(3)
  last_ele1% = last_ele1% + 1
  read_element_from_temp_item = 1
 ElseIf temp_item0(n%).poi(1) = -7 Then
  If read_element_from_item(temp_item0(n%).poi(0), temp_ele1(), last_temp_ele1%, _
        temp_ele2(), last_temp_ele2%) = 1 Then
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%).poi(0) = temp_item0(n%).poi(2)
        ele1(last_ele1%).poi(1) = temp_item0(n%).poi(3)
       last_ele1% = last_ele1% + 1
     For i% = 1 To last_temp_ele1% - 1
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%) = temp_ele1(i%)
        last_ele1% = last_ele1% + 1
     Next i%
     For i% = 1 To last_temp_ele2% - 1
      ReDim Preserve ele2(last_ele2%) As element_data_type
       ele2(last_ele2%) = temp_ele2(i%)
        last_ele2% = last_ele2% + 1
     Next i%
      read_element_from_temp_item = 1
  Else
   read_element_from_temp_item = 0
  End If
 ElseIf temp_item0(n%).poi(3) = -7 Then
  If read_element_from_item(temp_item0(n%).poi(2), temp_ele1(), last_temp_ele1%, _
        temp_ele2(), last_temp_ele2%) = 1 Then
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%).poi(0) = temp_item0(n%).poi(0)
        ele1(last_ele1%).poi(1) = temp_item0(n%).poi(1)
       last_ele1% = last_ele1% + 1
     For i% = 1 To last_temp_ele1% - 1
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%) = temp_ele1(i%)
        last_ele1% = last_ele1% + 1
     Next i%
     For i% = 1 To last_temp_ele2% - 1
      ReDim Preserve ele2(last_ele2%) As element_data_type
       ele2(last_ele2%) = temp_ele2(i%)
        last_ele2% = last_ele2% + 1
     Next i%
      read_element_from_temp_item = 1
  Else
   read_element_from_temp_item = 0
  End If
 Else
  If read_element_from_item(temp_item0(n%).poi(0), temp_ele1(), last_temp_ele1%, _
        temp_ele2(), last_temp_ele2%) = 1 Then
    If read_element_from_item(temp_item0(n%).poi(2), temp_ele1(), last_temp_ele1%, _
        temp_ele2(), last_temp_ele2%) = 1 Then
     For i% = 1 To last_temp_ele1% - 1
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%) = temp_ele1(i%)
        last_ele1% = last_ele1% + 1
     Next i%
     For i% = 1 To last_temp_ele2% - 1
      ReDim Preserve ele2(last_ele2%) As element_data_type
       ele1(last_ele2%) = temp_ele2(i%)
        last_ele2% = last_ele2% + 1
     Next i%
      read_element_from_temp_item = 1
  Else
   read_element_from_temp_item = 0
  End If
 Else
   read_element_from_temp_item = 0
  End If
 End If
ElseIf temp_item0(n%).sig = "/" Then
 If temp_item0(n%).poi(1) <> -7 And temp_item0(n%).poi(3) <> -7 Then
  ReDim Preserve ele1(last_ele1%) As element_data_type
  ele1(last_ele1%).poi(0) = temp_item0(n%).poi(0)
  ele1(last_ele1%).poi(1) = temp_item0(n%).poi(1)
  last_ele1% = last_ele1% + 1
  ReDim Preserve ele2(last_ele2%) As element_data_type
  ele2(last_ele1%).poi(0) = temp_item0(n%).poi(2)
  ele2(last_ele1%).poi(1) = temp_item0(n%).poi(3)
  last_ele2% = last_ele2% + 1
  read_element_from_temp_item = 1
 ElseIf temp_item0(n%).poi(1) = -7 Then
   If read_element_from_item(temp_item0(n%).poi(0), temp_ele1(), last_temp_ele1%, _
        temp_ele2(), last_temp_ele2%) = 1 Then
      ReDim Preserve ele2(last_ele2%) As element_data_type
       ele2(last_ele2%).poi(0) = temp_item0(n%).poi(2)
        ele2(last_ele2%).poi(1) = temp_item0(n%).poi(3)
       last_ele2% = last_ele2% + 1
     For i% = 1 To last_temp_ele1% - 1
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%) = temp_ele1(i%)
        last_ele1% = last_ele1% + 1
     Next i%
     For i% = 1 To last_temp_ele2% - 1
      ReDim Preserve ele1(last_ele2%) As element_data_type
       ele2(last_ele2%) = temp_ele2(i%)
        last_ele2% = last_ele2% + 1
     Next i%
      read_element_from_temp_item = 1
  Else
   read_element_from_temp_item = 0
  End If
 ElseIf temp_item0(n%).poi(3) = -7 Then
  If read_element_from_item(temp_item0(n%).poi(2), temp_ele1(), last_temp_ele1%, _
        temp_ele2(), last_temp_ele2%) = 1 Then
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%).poi(0) = temp_item0(n%).poi(0)
        ele1(last_ele1%).poi(1) = temp_item0(n%).poi(1)
       last_ele1% = last_ele1% + 1
     For i% = 1 To last_temp_ele1% - 1
      ReDim Preserve ele2(last_ele2%) As element_data_type
       ele2(last_ele2%) = temp_ele1(i%)
        last_ele2% = last_ele2% + 1
     Next i%
     For i% = 1 To last_temp_ele2% - 1
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%) = temp_ele1(i%)
        last_ele1% = last_ele1% + 1
     Next i%
      read_element_from_temp_item = 1
  Else
   read_element_from_temp_item = 0
  End If
 Else
   If read_element_from_item(temp_item0(n%).poi(0), temp_ele1(), last_temp_ele1%, _
        temp_ele2(), last_temp_ele2%) = 1 Then
    If read_element_from_item(temp_item0(n%).poi(2), temp_ele2(), last_temp_ele2%, _
        temp_ele1(), last_temp_ele1%) = 1 Then
     For i% = 1 To last_temp_ele1% - 1
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%) = temp_ele1(i%)
        last_ele1% = last_ele1% + 1
     Next i%
     For i% = 1 To last_temp_ele2% - 1
      ReDim Preserve ele2(last_ele2%) As element_data_type
       ele2(last_ele2%) = temp_ele2(i%)
        last_ele2% = last_ele2% + 1
     Next i%
      read_element_from_temp_item = 1
  Else
   read_element_from_temp_item = 0
  End If
 Else
   read_element_from_temp_item = 0
  End If
 End If
Else
  read_element_from_temp_item = 0
End If
End Function
Public Function from_element_to_item0(ele() As element_data_type, last_ele%) As Integer
Dim tem_conditions As condition_data_type
Dim tn%
If last_ele% = 2 Then
 Call set_item0(ele(0).poi(0), ele(0).poi(1), ele(1).poi(0), ele(1).poi(1), _
           "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, tem_conditions, 0, _
             from_element_to_item0, 0, 0, condition_data0, False)
Else
 tn% = from_element_to_item0(ele(), last_ele% - 1)
 Call set_item0(ele(last_ele% - 1).poi(0), ele(last_ele% - 1).poi(1), tn%, -7, _
           "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, tem_conditions, 0, _
              from_element_to_item0, 0, 0, condition_data0, False)
End If
End Function

Public Sub draw_any_triangle(w As wentitype)
Dim i%
For i% = 0 To 2
If draw_free_point(w.point_no(i%), _
     w.condition(i%)) Then
  Exit Sub
End If
    Call change_point_degree(w.point_no(i%), -3)
Next i%
Call draw_triangle(w.point_no(0), w.point_no(1), w.point_no(2), condition)
End Sub
Public Sub draw_triangle(ByVal p1%, ByVal p2%, ByVal p3%, condition_or_conclusion As Byte)
Dim color As Byte
If condition_or_conclusion = condition Then
   color = condition_color
Else
   color = conclusion_color
End If
 If condition_or_conclusion = conclusion Then
 Call line_number(p1%, p2%, pointapi0, pointapi0, _
                  depend_condition(0, 0), depend_condition(0, 0), _
                  condition_or_conclusion, color, 1, 0)
 Call line_number(p2%, p3%, pointapi0, pointapi0, _
                  depend_condition(0, 0), depend_condition(0, 0), _
                  condition_or_conclusion, color, 1, 0)
 Call line_number(p1%, p3%, pointapi0, pointapi0, _
                  depend_condition(0, 0), depend_condition(0, 0), _
                  condition_or_conclusion, color, 1, 0)
 Else
 Call line_number(p1%, p2%, pointapi0, pointapi0, _
                  depend_condition(point_, p1%), depend_condition(point_, p2%), _
                  condition_or_conclusion, color, 1, 0)
 Call line_number(p2%, p3%, pointapi0, pointapi0, _
                  depend_condition(point_, p2%), depend_condition(point_, p3%), _
                  condition_or_conclusion, color, 1, 0)
 Call line_number(p1%, p3%, pointapi0, pointapi0, _
                  depend_condition(point_, p1%), depend_condition(point_, p3%), _
                  condition_or_conclusion, color, 1, 0)
 End If
End Sub
Public Sub draw_any_polygon4(w As wentitype)
Dim i%, t_line%
 For i% = 0 To 3
If draw_free_point(w.point_no(i%), _
      w.condition(i%)) Then
      Exit Sub
End If
    Call change_point_degree(w.point_no(i%), -3)
Next i%
Call draw_polygon4(w.point_no(0), w.point_no(1), w.point_no(2), w.point_no(3), condition)
End Sub
Public Sub draw_polygon4(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, condition_or_conclusion As Byte)
Dim color As Byte
If condition_or_conclusion = condition Then
   color = condition_color
Else
   color = conclusion_color
End If
 If condition_or_conclusion = conclusion Then
 Call line_number(p1%, p2%, pointapi0, pointapi0, _
                  depend_condition(0, 0), depend_condition(0, 0), _
                  condition_or_conclusion, color, 1, 0)
 Call line_number(p2%, p3%, pointapi0, pointapi0, _
                  depend_condition(0, 0), depend_condition(0, 0), _
                  condition_or_conclusion, color, 1, 0)
 Call line_number(p3%, p4%, pointapi0, pointapi0, _
                  depend_condition(0, 0), depend_condition(0, 0), _
                  condition_or_conclusion, color, 1, 0)
 Call line_number(p1%, p4%, pointapi0, pointapi0, _
                  depend_condition(0, 0), depend_condition(0, 0), _
                  condition_or_conclusion, color, 1, 0)
 Else
 Call line_number(p1%, p2%, pointapi0, pointapi0, _
                  depend_condition(point_, p1%), depend_condition(point_, p2%), _
                  condition_or_conclusion, color, 1, 0)
 Call line_number(p2%, p3%, pointapi0, pointapi0, _
                  depend_condition(point_, p2%), depend_condition(point_, p3%), _
                  condition_or_conclusion, color, 1, 0)
 Call line_number(p4%, p3%, pointapi0, pointapi0, _
                  depend_condition(point_, p4%), depend_condition(point_, p3%), _
                  condition_or_conclusion, color, 1, 0)
 Call line_number(p1%, p4%, pointapi0, pointapi0, _
                  depend_condition(point_, p1%), depend_condition(point_, p4%), _
                  condition_or_conclusion, color, 1, 0)
 End If
End Sub
Public Function input_data_from_item(it As item0_data_type, val As String, re As total_record_type) As Byte
Dim temp_record As total_record_type
Dim tri_f As tri_function_data_type
temp_record = re
If it.poi(1) > 0 Then
 input_data_from_item = set_line_value(it.poi(0), it.poi(1), val, 0, 0, 0, temp_record.record_data, 0, 0, False)
  If input_data_from_item > 1 Then
   Exit Function
  End If

ElseIf it.poi(1) = -7 Then
 'input_data_from_item = set_item
ElseIf it.poi(1) = -1 Then
 input_data_from_item = set_tri_function(it.poi(0), val, "", "", "", 0, temp_record, False, tri_f, 0)
  If input_data_from_item > 1 Then
   Exit Function
  End If
ElseIf it.poi(1) = -2 Then
 input_data_from_item = set_tri_function(it.poi(0), "", val, "", "", 0, temp_record, False, tri_f, 0)
  If input_data_from_item > 1 Then
   Exit Function
  End If
ElseIf it.poi(1) = -3 Then
 input_data_from_item = set_tri_function(it.poi(0), "", "", val, "", 0, temp_record, False, tri_f, 0)
  If input_data_from_item > 1 Then
   Exit Function
  End If
ElseIf it.poi(1) = -4 Then
 input_data_from_item = set_tri_function(it.poi(0), "", "", "", val, 0, temp_record, False, tri_f, 0)
  If input_data_from_item > 1 Then
   Exit Function
  End If
ElseIf it.poi(1) = -6 Then
 input_data_from_item = set_angle_value(it.poi(0), val, temp_record, 0, 0, False)
  If input_data_from_item > 1 Then
   Exit Function
  End If
End If
End Function

Public Function read_element_for_string(ByVal s$, cond_ty As Byte, no%, para1$, para2$) As Boolean
Dim i%, j%, k%, m%, n%
Dim ts(4) As String
Dim ty As Byte
Dim v As String
Dim p$
For i% = 1 To last_conditions.last_cond(1).line_value_no
 m% = InStr(1, line_value(i%).data(0).data0.value, s$, 0)
   If m% > 0 Then
    v = solve_first_order_equation(angle3_value(i%).data(0).data0.value, "l", s$)
     ty = string_type(v, ts(0), ts(1), ts(2), ts(3))
      If ty = 3 Then
      Else
      End If
   End If
Next i%
End Function

Public Function read_string_value_for_condition(ByVal s$, ty As Byte, n%) As Boolean
Dim i%
For i% = 1 To last_conditions.last_cond(1).line_value_no
 If minus_string(line_value(i%).data(0).data0.value, s$, True, False) = "0" Then
  read_string_value_for_condition = True
   ty = line_value_
    n% = i%
     Exit Function
 End If
Next i%
For i% = 1 To last_conditions.last_cond(1).angle3_value_no
 If angle3_value(i%).data(0).data0.angle(1) = 0 Then
     If minus_string(angle3_value(i%).data(0).data0.value, s$, True, False) = "0" Then
       read_string_value_for_condition = True
        ty = angle3_value_
         n% = i%
     Exit Function
     End If
 End If
Next i%
For i% = 1 To last_conditions.last_cond(1).relation_no
 If minus_string(Drelation(i%).data(0).data0.value, s$, True, False) = "0" Then
       read_string_value_for_condition = True
        ty = relation_
         n% = i%
     Exit Function
 ElseIf minus_string(Drelation(i%).data(0).data0.value, divide_string("1", s$, False, False), True, False) = "0" Then
       read_string_value_for_condition = True
        ty = relation_
         n% = -i%
     Exit Function
    
 End If
Next i%
End Function

Public Sub set_draw_width()
Call draw_again1(Draw_form)
line_width = temp_line_width
condition_color = temp_condition_color
conclusion_color = temp_conclusion_color
fill_color = temp_fill_color
regist_data.line_width = line_width
regist_data.conclusion_color = conclusion_color
regist_data.condition_color = condition_color
regist_data.fill_color = fill_color
Draw_form.DrawWidth = line_width
Call init_set
Call init_color
'widthform.Hide
Call draw_again1(Draw_form)

End Sub

Public Sub read_multi_item_(ByVal num%, ByVal st%, ByVal en%, ByVal para0$, para() As String, item() As item0_data_type, last%)
Dim i%, brace_n%, sig_n%, last1%
Dim Lbrace%, Rbrace%
Dim sig_(8) As Integer
Dim sig(8) As String
Dim tsig(8) As String
Dim tsig_(8) As Integer
Dim tsig_no As Integer
Dim t_item As item0_data_type
Dim t_item_(8) As item0_data_type
Dim t_item1_(8) As item0_data_type
Dim t_para As String
Dim t_para_(8) As String
Dim t_para1_(8) As String
Dim brace_no%
Lbrace% = -1
If st% >= en% Then
 last% = 0
Else
For i% = st% To en%
If i% > st% And (C_display_wenti.m_condition(num, i%) = "+" Or C_display_wenti.m_condition(num, i%) = "-") Then
   sig_(sig_n%) = i%
    sig(sig_n%) = C_display_wenti.m_condition(num, i%)
     sig_n% = sig_n% + 1
ElseIf C_display_wenti.m_condition(num, i%) = "(" Then
  If brace_n% = 0 And (Lbrace% = -1 Or brace_no% > 0) Then
  Lbrace% = i%
  End If
  brace_n% = brace_n% + 1
  brace_no% = brace_no% + 1
ElseIf C_display_wenti.m_condition(num, i%) = ")" Then
  brace_n% = brace_n% - 1
  If brace_n% = 0 Then
  Rbrace% = i%
  End If
ElseIf C_display_wenti.m_condition(num, i%) = "*" Or C_display_wenti.m_condition(num, i%) = "/" Then
 tsig_no = tsig_no + 1
 tsig(tsig_no) = C_display_wenti.m_condition(num, i%)
 tsig_(tsig_no) = i%
ElseIf Asc(C_display_wenti.m_condition(num, i%)) = 13 Or _
  C_display_wenti.m_condition(num, i%) = empty_char Then
   en% = i - 1
End If
Next i%
If Rbrace% > 0 And Rbrace < en% Then
 If C_display_wenti.m_condition(num, Rbrace% + 1) = "/" Then
    Call read_multi_item(num%, Rbrace% + 2, en%, t_item, t_para, 1, 1)
     If t_item.sig = "F" Then
       If para0$ = "" Then
          para0$ = divide_string("1", t_para, True, False)
       Else
          para0$ = divide_string(para0$, t_para, True, False)
       End If
      Call read_multi_item_(num%, st%, Rbrace%, para0$, para(), item(), last%)
      Exit Sub
     Else
     End If
 ElseIf C_display_wenti.m_condition(num, Rbrace% + 1) = "*" Then
    Call read_multi_item(num%, Rbrace% + 2, en%, t_item, t_para, 1, 0)
     If t_item.sig = "F" Then
       If para0$ = "" Then
          para0$ = t_para
       Else
          para0$ = time_string(para0$, t_para, True, False)
       End If
      Call read_multi_item_(num%, st%, Rbrace%, para0$, para(), item(), last%)
      Exit Sub
     Else
     End If
 Else
    Call read_multi_item(num%, Rbrace% + 1, en%, t_item, t_para, 1, 1)
      If t_item.sig = "F" Then
       If para0$ = "" Then
          para0$ = t_para
       Else
          para0$ = time_string(para0$, t_para, True, False)
       End If
      Call read_multi_item_(num%, st%, Rbrace%, para0$, para(), item(), last%)
      Exit Sub
     Else
     End If
 End If
ElseIf tsig_no > 1 Then
 If tsig(tsig_no% - 1) = "*" Then
    Call read_multi_item(num%, tsig_(tsig_no% - 1) + 1, en%, t_item, t_para, 1, 1)
     If t_item.sig = "F" Then
       If para0$ = "" Then
          para0$ = t_para
       Else
          para0$ = time_string(para0$, t_para, True, False)
       End If
      Call read_multi_item_(num%, st%, tsig_(tsig_no% - 1) - 1, para0$, para(), item(), last%)
      Exit Sub
     Else
     End If
 Else
  
 End If
End If
If Lbrace% = -1 Or (Lbrace% > sig_(0) And sig_(0) > 0) Then
   If C_display_wenti.m_condition(num, st%) = "+" Then
    If sig_(0) = 0 Then
     Call read_multi_item(num%, st% + 1, en%, t_item, t_para, 1, 0)
      last% = 1
       para(0) = time_string(para0$, t_para, True, False)
        item(0) = t_item
    Else
      Call read_multi_item(num%, st% + 1, sig_(0) - 1, t_item, t_para, 1, 0)
       Call read_multi_item_(num%, sig_(0), en%, para0, t_para_(), t_item_(), last%)
         para(0) = time_string(para0$, t_para, True, False)
          item(0) = t_item
          For i% = 0 To last% - 1
           para(1 + i%) = t_para_(i%)
           item(1 + i%) = t_item_(i%)
          Next i%
            last% = last% + 1
    End If
   ElseIf C_display_wenti.m_condition(num, st%) = "-" Then
    If sig_(0) = 0 Then
     Call read_multi_item(num%, st% + 1, en%, t_item, t_para, 1, 0)
      last% = 1
       para(0) = time_string(para0$, t_para, True, False)
        para(0) = time_string("-1", para(0), True, False)
        item(0) = t_item
    Else
      Call read_multi_item(num%, st% + 1, sig_(0) - 1, t_item, t_para, 1, 0)
       Call read_multi_item_(num%, sig_(0), en%, para0, t_para_(), t_item_(), last%)
         para(0) = time_string(para0$, t_para, True, False)
          para(0) = time_string(para0$, t_para, True, False)
           item(0) = t_item
          For i% = 0 To last% - 1
           para(1 + i%) = t_para_(i%)
           item(1 + i%) = t_item_(i%)
          Next i%
            last% = last% + 1
    End If
   Else
    If sig_(0) = 0 Then
     Call read_multi_item(num%, st%, en%, t_item, t_para, 1, 0)
      last% = 1
       para(0) = time_string(para0$, t_para, True, False)
        item(0) = t_item
    Else
      Call read_multi_item(num%, st%, sig_(0) - 1, t_item, t_para, 1, 0)
       Call read_multi_item_(num%, sig_(0), en%, para0, t_para_(), t_item_(), last%)
         para(0) = time_string(para0$, t_para, True, False)
          item(0) = t_item
          For i% = 0 To last% - 1
           para(1 + i%) = t_para_(i%)
           item(1 + i%) = t_item_(i%)
          Next i%
            last% = last% + 1
    End If
   End If
Else
 If Lbrace% > st% Then
 If C_display_wenti.m_condition(num, Lbrace% - 1) = "*" Then
  'Call read_number_from_wenti(num, st%, Lbrace% - 2, t_para)
 Else
  'Call read_number_from_wenti(num, st%, Lbrace% - 1, t_para)
 End If
 Else
  t_para = "1"
 End If
  If t_para = "" Or t_para = "+" Then
     t_para = "1"
  ElseIf t_para = "-" Then
    t_para = "-1"
  End If
    Call read_multi_item_(num%, Lbrace% + 1, Rbrace% - 1, _
           time_string(para0, t_para, True, False), t_para_(), t_item_(), last1%)
    Call read_multi_item_(num%, Rbrace% + 1, en%, para0, t_para1_(), t_item1_(), last%)
    For i% = 0 To last1% - 1
     item(i%) = t_item_(i%)
     para(i%) = t_para_(i%)
    Next i%
    For i% = 0 To last% - 1
     item(last1% + i%) = t_item_(i%)
     para(last1% + i%) = t_para_(i%)
    Next i%
    last% = last% + last1%
 End If
 End If
End Sub

Private Function change_picture_54_23_22(wenti_data As wentitype, change_element As condition_type) As Boolean
'-54□□的垂直平分线交□□于□
'-22 过□点平行□□的直线交□□于□
'-23过□点垂直□□的直线交□□于□
Dim is_ch As Boolean
Dim mid_coord As POINTAPI
Dim paral_or_verti_ As Integer
Dim t_cond As condition_data_type
If wenti_data.no = -22 Then
 paral_or_verti_ = paral_
ElseIf wenti_data.no = -23 Then
 paral_or_verti_ = verti_
ElseIf wenti_data.no = -54 Then
          paral_or_verti_ = verti_
   If (change_element.ty = point_ And change_element.no = wenti_data.poi(3)) Or _
                (change_element.ty = point_ And change_element.no = wenti_data.poi(4)) Then
     m_poi(wenti_data.poi(2)).data(0).data0.coordinate = mid_POINTAPI( _
      m_poi(wenti_data.poi(3)).data(0).data0.coordinate, m_poi(wenti_data.poi(4)).data(0).data0.coordinate)
     'm_poi(wenti_data.poi(2)).data(0).is_change = True
     Call change_m_point(wenti_data.poi(2))
    End If
End If
  'If m_poi(wenti_data.poi(2)).data(0).is_change Or m_lin(wenti_data.line_no(2)).data(0).is_change Then
  '   Call change_paral_or_verti_line(wenti_data.line_no(3), wenti_data.poi(2), wenti_data.line_no(2), paral_or_verti_)
 ' End If
      Call inter_point_line_line3(0, 0, wenti_data.line_no(1), wenti_data.poi(2), paral_or_verti_, _
             wenti_data.line_no(3), m_poi(wenti_data.poi(1)).data(0).data0.coordinate, _
              0, True, t_cond, False)
      'm_poi(wenti_data.poi(1)).data(0).is_change = True
      Call change_m_point(wenti_data.poi(1))
End Function



Private Function change_picture_50(wenti_data As wentitype, change_element As condition_type) As Boolean
Dim coord(1) As POINTAPI
Dim ty(1) As Integer
Dim t_cp%
       If m_lin(wenti_data.line_no(3)).data(0).is_change = 255 Or _
           m_lin(wenti_data.line_no(4)).data(0).is_change = 255 Then
          Call draw_equal_angle(wenti_data.poi(1), wenti_data.poi(2), wenti_data.poi(3), wenti_data.poi(4))
         ' m_poi(wenti_data.poi(4)).data(0).is_change = True
         ' Call change_m_point(wenti_data.poi(4))
       End If
End Function
Private Function change_picture_43(wenti_data As wentitype) As Boolean
'-43在□□上取一点□使得□□＝!_~
Dim p_coord As POINTAPI
Dim r&
change_picture_43 = True
If wenti_data.poi(1) > yidian_no Then
If m_lin(wenti_data.line_no(1)).data(0).is_change = 255 Or _
    m_Circ(wenti_data.circ(2)).data(0).is_change Then
 If inter_point_line_circle1( _
    m_poi(m_lin(wenti_data.line_no(1)).data(0).data0.poi(0)).data(0).data0.coordinate, _
       paral_, m_lin(wenti_data.line_no(1)).data(0), _
        m_Circ(wenti_data.circ(2)).data(0).data0, _
         t_coord1, 0, t_coord2, 0) > 0 Then
 If wenti_data.inter_set_point_type = 1 Then
             Call set_point_coordinate(wenti_data.poi(1), t_coord1, True)
 Else
             Call set_point_coordinate(wenti_data.poi(1), t_coord2, True)
 End If
 Else
   change_picture_43 = False
 End If
End If
End If
End Function

Private Function change_picture_42_57(wenti_data As wentitype) As Boolean
'-42 在⊙□[down\\(_)]上取一点□使得□□＝!_~
'-57 在⊙□□□上取一点□使得□□＝!_~
change_picture_42_57 = True
If wenti_data.poi(1) > yidian_no Then
 If m_Circ(wenti_data.circ(1)).data(0).is_change Or _
     m_Circ(wenti_data.circ(2)).data(0).is_change Then
  If inter_point_circle_circle_(m_Circ(wenti_data.circ(1)).data(0).data0, _
      m_Circ(wenti_data.circ(2)).data(0).data0, _
        t_coord1, 0, t_coord2, 0, 0, 0, True) > 0 Then
   If wenti_data.inter_set_point_type = 1 Then
            Call set_point_coordinate(wenti_data.poi(1), t_coord1, True)
   Else
            Call set_point_coordinate(wenti_data.poi(1), t_coord2, True)
   End If
  Else
  change_picture_42_57 = False
  End If
 End If
 End If
End Function

Private Function change_picture_33_44(wenti_data As wentitype) As Boolean
'-33 过□作⊙□[down\\(_)]的切线□□
'-44 过□作⊙□□□的切线□□
Dim j%
Dim t_coord(1) As POINTAPI
Dim ty As Boolean
If wenti_data.poi(1) >= yidian_no Then
change_picture_33_44 = True
If m_Circ(wenti_data.circ(1)).data(0).is_change Or _
     m_poi(wenti_data.poi(1)).data(0).is_change Then
     If wenti_data.inter_set_point_type = tangent_line_by_point_on_circle Then
     m_lin(wenti_data.line_no(1)).data(0).data0.depend_poi1_coord = add_POINTAPI( _
        m_poi(wenti_data.poi(1)).data(0).data0.coordinate, _
         verti_POINTAPI(minus_POINTAPI( _
          m_Circ(wenti_data.circ(1)).data(0).data0.c_coord, _
             m_poi(wenti_data.poi(1)).data(0).data0.coordinate)))
        'm_lin(wenti_data.line_no(1)).data(0).is_change = True
      Call change_m_line(wenti_data.line_no(1))
     Else 'If wenti_data.inter_set_point_type = 1 Then
       'm_poi(wenti_data.poi(2)).data(0).data0.coordinate = _
         inter_point_circle_circle_by_pointapi(m_Circ(wenti_data.circ(1)).data(0).data0.c_coord, _
           m_Circ(wenti_data.circ(1)).data(0).data0.radii, _
             mid_POINTAPI(m_poi(wenti_data.poi(1)).data(0).data0.coordinate, _
               m_Circ(wenti_data.circ(1)).data(0).data0.c_coord), _
                distance_of_two_POINTAPI(m_poi(wenti_data.poi(1)).data(0).data0.coordinate, _
                 m_Circ(wenti_data.circ(1)).data(0).data0.c_coord) / 2, _
                t_coord(0), t_coord(1), wenti_data.inter_set_point_type)
           ' m_poi(wenti_data.poi(2)).data(0).is_change = True
           ' Call change_m_point(wenti_data.poi(2))
     End If
'ElseIf m_poi(wenti_data.poi(1)).data(0).is_change Then
'  t_coord = minus_POINTAPI( _
         m_poi(wenti_data.poi(1)).data(0).data0.coordinate, _
            m_poi(wenti_data.poi(2)).data(0).data0.coordinate)
'  If m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.X - _
            m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X <> 0 Then
'    A! = -(m_poi(wenti_data.poi(1)).data(0).data0.coordinate.Y - _
         m_poi(wenti_data.poi(2)).data(0).data0.coordinate.Y) / _
          (m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.X - _
            m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X)
'  ElseIf m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.Y - _
            m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y <> 0 Then
'    A! = (m_poi(wenti_data.poi(1)).data(0).data0.coordinate.X - _
         m_poi(wenti_data.poi(2)).data(0).data0.coordinate.X) / _
          (m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.Y - _
            m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y)
'  End If
'      Call C_display_wenti.set_m_point_no(wenti_data.wenti_no, Int(A! * 1000), 7, False)
End If

'If wenti_data.inter_set_point_type = 0 Then
't_coord = minus_POINTAPI(time_POINTAPI_by_number( _
'             m_poi(wenti_data.poi(2)).data(0).data0.coordinate, 2), _
              m_poi(wenti_data.poi(1)).data(0).data0.coordinate)
 '2 * m_poi(C_display_wenti.m_point_no(45)).data(0).data0.coordinate.X - _
  m_poi(C_display_wenti.m_point_no(44)).data(0).data0.coordinate.X
't_coord.Y = _
 2 * m_poi(C_display_wenti.m_point_no(45)).data(0).data0.coordinate.Y - _
  m_poi(C_display_wenti.m_point_no(44)).data(0).data0.coordinate.Y
' Else
't_coord = minus_POINTAPI(time_POINTAPI_by_number( _
              m_poi(wenti_data.poi(1)).data(0).data0.coordinate, 2), _
               m_poi(wenti_data.poi(2)).data(0).data0.coordinate)
't_coord.X = _
 2 * m_poi(C_display_wenti.m_point_no(44)).data(0).data0.coordinate.X - _
  m_poi(C_display_wenti.m_point_no(45)).data(0).data0.coordinate.X
't_coord.Y = _
 2 * m_poi(C_display_wenti.m_point_no(44)).data(0).data0.coordinate.Y - _
  m_poi(C_display_wenti.m_point_no(45)).data(0).data0.coordinate.Y
' End If
'   Call set_point_coordinate(wenti_data.poi(3), t_coord, True)
End If
End Function

Private Function change_picture_32(wenti_data As wentitype) As Boolean
'与⊙□[down\\(_)]相切于点□的切线交直线□□于□
Dim t_cond As condition_data_type
If wenti_data.poi(1) > yidian_no Then
change_picture_32 = True
If m_Circ(wenti_data.circ(1)).data(0).is_change Or _
     m_lin(wenti_data.line_no(1)).data(0).is_change = 255 Or _
      m_poi(wenti_data.poi(2)).data(0).is_change Then
Call inter_point_line_line3(wenti_data.poi(2), False, wenti_data.line_no(3), _
        m_lin(wenti_data.line_no(1)).data(0).data0.poi(0), _
         True, wenti_data.line_no(1), t_coord, wenti_data.poi(1), True, t_cond, False)
 t_coord = minus_POINTAPI(time_POINTAPI_by_number( _
        m_poi(wenti_data.poi(2)).data(0).data0.coordinate, 2), _
         m_poi(wenti_data.poi(1)).data(0).data0.coordinate)
 't_coord.Y = 2 * m_poi(wenti_data.m_point_no(39)).data(0).data0.coordinate.Y - _
       m_poi(wenti_data.m_point_no(44)).data(0).data0.coordinate.Y
       Call set_point_coordinate(wenti_data.poi(3), t_coord, True)
End If
End If
End Function

Private Function change_picture_31_30(wenti_data As wentitype) As Boolean
'-30在⊙□[down\\(_)]上取一点□使得□□＝□□
'-58在⊙□□□上取一点□使得□□＝□□
'-31在□□上取一点□使得□□＝□□
change_picture_31_30 = True
If wenti_data.poi(1) > yidian_no Then
If wenti_data.no = -31 Then
  If m_lin(wenti_data.line_no(1)).data(0).is_change = 255 Or _
       m_Circ(wenti_data.circ(1)).data(0).is_change Then
   Call inter_point_line_circle3( _
   m_poi(wenti_data.point_no(0)).data(0).data0.coordinate, _
    True, m_poi(wenti_data.point_no(1)).data(0).data0.coordinate, _
      m_poi(wenti_data.point_no(0)).data(0).data0.coordinate, _
       m_Circ(wenti_data.circ(1)).data(0).data0, _
        t_coord1, 0, t_coord2, 0, wenti_data.inter_set_point_type, True)
   If wenti_data.inter_set_point_type = 2 Then
     Call set_point_coordinate(wenti_data.poi(1), t_coord2, True)
   Else
     Call set_point_coordinate(wenti_data.poi(1), t_coord1, True)
  End If
 End If
Else '30
 If m_Circ(wenti_data.circ(1)).data(0).is_change Or _
              m_Circ(wenti_data.circ(2)).data(0).is_change Then
   Call inter_point_circle_circle_( _
     m_Circ(wenti_data.circ(1)).data(0).data0, _
      m_Circ(wenti_data.circ(2)).data(0).data0, _
       t_coord1, 0, t_coord2, 0, 0, 0, True)
  If wenti_data.inter_set_point_type = 1 Then
    Call set_point_coordinate(wenti_data.poi(1), t_coord2, True)
  Else
    Call set_point_coordinate(wenti_data.poi(1), t_coord1, True)
  End If
 End If
End If
End If
End Function

Private Function change_picture_15(wenti_data As wentitype) As Boolean
'-15 □□□□是梯形
change_picture_15 = True
 If wenti_data.point_no(4) = 1 Then
  If m_poi(wenti_data.point_no(0)).data(0).is_change Or _
    m_poi(wenti_data.point_no(1)).data(0).is_change Or _
      m_poi(wenti_data.point_no(2)).data(0).is_change Then
   m_poi(wenti_data.point_no(3)).data(0).data0.coordinate.X = _
       m_poi(wenti_data.point_no(2)).data(0).data0.coordinate.X + _
        (m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X - _
          m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.X) * _
          wenti_data.point_no(5) / 1000
      m_poi(wenti_data.point_no(3)).data(0).data0.coordinate.Y = _
       m_poi(wenti_data.point_no(2)).data(0).data0.coordinate.Y + _
        (m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y - _
          m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.Y) * _
          wenti_data.point_no(5) / 1000
        'poi(wenti_data.point_no(3)).data(0).is_change
        ' Call change_circle_(wenti_data.point_no(3), pointapi0)
      If Abs(m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.X - _
           m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X) > 5 Then
           Call C_display_wenti.set_m_point_no(wenti_data.wenti_no, _
           (m_poi(wenti_data.point_no(3)).data(0).data0.coordinate.X - _
            m_poi(wenti_data.point_no(2)).data(0).data0.coordinate.X) * 1000 / _
              (m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X - _
                m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.X), 5, False)
      Else
           Call C_display_wenti.set_m_point_no(wenti_data.wenti_no, _
           (m_poi(wenti_data.point_no(3)).data(0).data0.coordinate.Y - _
            m_poi(wenti_data.point_no(2)).data(0).data0.coordinate.Y) * 1000 / _
              (m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y - _
                m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.Y), 5, False)
      End If
   End If
 End If

End Function

Private Function change_picture_14_13_11_10(wenti_data As wentitype) As Boolean
'-10 □□□□是菱形
'-11 □□□□是平行四边形
'-13 □□□□是长方形
'-14 □□□□是等腰梯形
Dim change_point%
Dim dr_ty As Integer
change_picture_14_13_11_10 = True
  If m_poi(wenti_data.point_no(0)).data(0).is_change Or _
      m_poi(wenti_data.point_no(1)).data(0).is_change Or _
       m_poi(wenti_data.point_no(2)).data(0).is_change Or _
        m_poi(wenti_data.point_no(3)).data(0).is_change Then
        temp_four_point_fig.poi(0) = wenti_data.point_no(0)
        temp_four_point_fig.poi(1) = wenti_data.point_no(1)
        temp_four_point_fig.poi(2) = wenti_data.point_no(2)
        temp_four_point_fig.poi(3) = wenti_data.point_no(3)
        temp_four_point_fig.p(0) = _
             m_poi(wenti_data.point_no(0)).data(0).data0.coordinate
        temp_four_point_fig.p(1) = _
             m_poi(wenti_data.point_no(1)).data(0).data0.coordinate
        temp_four_point_fig.p(2) = _
             m_poi(wenti_data.point_no(2)).data(0).data0.coordinate
        temp_four_point_fig.p(3) = _
             m_poi(wenti_data.point_no(3)).data(0).data0.coordinate
        If m_poi(wenti_data.point_no(2)).data(0).is_change Then
           change_point% = wenti_data.point_no(2)
           dr_ty = 1
        ElseIf m_poi(wenti_data.point_no(3)).data(0).is_change Then
           change_point% = wenti_data.point_no(3)
           dr_ty = 2
        Else
           change_point% = wenti_data.point_no(2)
           dr_ty = 0
        End If

 If wenti_data.no = -14 Then
   Call set_temp_equal_side_tixing(m_poi(change_point%).data(0).data0.coordinate, dr_ty)
 ElseIf wenti_data.no = -13 Then
   Call set_temp_long_squre(m_poi(change_point%).data(0).data0.coordinate, dr_ty)
 ElseIf wenti_data.no = -11 Then
   Call set_temp_parallelogram(m_poi(change_point%).data(0).data0.coordinate, dr_ty)
 ElseIf wenti_data.no = -10 Then
   Call set_temp_rhombus(m_poi(change_point%).data(0).data0.coordinate, dr_ty)
 End If
 End If
End Function

Private Function change_picture_3(wenti_data As wentitype) As Boolean
'-3 与⊙□[down\\(_)]相切于点□的切线交⊙□[down\\(_)]
If wenti_data.poi(1) > yidian_no Then
change_picture_3 = True
If m_Circ(wenti_data.circ(1)).data(0).is_change Or _
    m_Circ(wenti_data.circ(2)).data(0).is_change Or _
     m_poi(wenti_data.poi(2)).data(0).is_change Then
Call inter_point_line_circle3( _
m_poi(wenti_data.poi(2)).data(0).data0.coordinate, _
 False, m_poi(wenti_data.poi(2)).data(0).data0.coordinate, _
  m_poi(m_Circ(wenti_data.circ(1)).data(0).data0.center).data(0).data0.coordinate, _
    m_Circ(wenti_data.circ(2)).data(0).data0, _
      t_coord1, 0, t_coord2, 0, wenti_data.inter_set_point_type, True)
If wenti_data.inter_set_point_type = 2 Then
 Call set_point_coordinate(wenti_data.poi(1), t_coord1, True)
Else
 Call set_point_coordinate(wenti_data.poi(1), t_coord2, True)
End If
 t_coord = minus_POINTAPI(time_POINTAPI_by_number( _
      m_poi(wenti_data.poi(2)).data(0).data0.coordinate, 2), _
       m_poi(wenti_data.poi(1)).data(0).data0.coordinate)
Call set_point_coordinate(wenti_data.poi(3), t_coord, True)
End If
End If
End Function


Private Function change_picture_18(wenti_data As wentitype) As Boolean
'-18 △□□□是等腰三角形
Dim b!
Dim c!
change_picture_18 = True
If m_poi(wenti_data.point_no(0)).data(0).is_change Or _
     m_poi(wenti_data.point_no(1)).data(0).is_change Then
b! = sqr((m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X - m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.X) ^ 2 + _
        (m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y - m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.Y) ^ 2)
c! = sqr((m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X - m_poi(wenti_data.point_no(2)).data(0).data0.coordinate.X) ^ 2 + _
        (m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y - m_poi(wenti_data.point_no(2)).data(0).data0.coordinate.Y) ^ 2)
'Call draw_point(Draw_form, m_poi(wenti_data.point_no(2)).data(0).data0, delete)
m_poi(wenti_data.point_no(2)).data(0).data0.coordinate.X = m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X + _
 (m_poi(wenti_data.point_no(2)).data(0).data0.coordinate.X - m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X) * b! / c!
m_poi(wenti_data.point_no(2)).data(0).data0.coordinate.Y = m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y + _
 (m_poi(wenti_data.point_no(2)).data(0).data0.coordinate.Y - m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y) * b! / c!
       'poi(wenti_data.point_no(2)).data(0).is_change
       '  Call change_circle_(wenti_data.point_no(2), pointapi0)
End If
End Function

Private Function change_picture_16_12_9_8(wenti_data As wentitype, change_element As condition_type) As Boolean
'-8 □□□□□□是正六边形
'-9 □□□□□是正五边形
'-12 □□□□是正方形
'-16 △□□□是等边三角形
Dim i%, j%
change_picture_16_12_9_8 = True
 If wenti_data.no = -16 Then
  j% = 3
 ElseIf wenti_data.no = -12 Then
  j% = 4
 ElseIf wenti_data.no = -9 Then
  j% = 5
 ElseIf wenti_data.no = -8 Then
  j% = 6
End If
If (change_element.ty = point_ And change_element.no = wenti_data.point_no(0)) Or _
      (change_element.ty = point_ And change_element.no = wenti_data.point_no(1)) Then
Call set_polygon2(wenti_data.point_no(0), _
                  wenti_data.point_no(1), _
                   j%, poly(wenti_data.point_no(10)), True)
 For i% = 2 To j% - 1
  'poi(wenti_data.point_no(i%)).data(0).is_change
  '  Call change_circle_(wenti_data.point_no(i%), pointapi0)
Next i%
 End If

End Function

Private Function change_picture_17(wenti_data As wentitype) As Boolean
'-17 △□□□是等腰直角三角形
Dim u%, v%
change_picture_17 = True
If m_poi(wenti_data.point_no(0)).data(0).is_change Or _
    m_poi(wenti_data.point_no(1)).data(0).is_change Then
u% = m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.X - _
     m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X
v% = m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.Y - _
     m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y
If wenti_data.point_no(10) = 0 Then
 t_coord.X = m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X _
    - v%
 t_coord.Y = m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y _
    + u%
    Call set_point_coordinate(wenti_data.point_no(2), t_coord, True)
 ElseIf wenti_data.point_no(10) = 1 Then
  t_coord.X = m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X _
    + v%
  t_coord.Y = m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y _
    - u%
    Call set_point_coordinate(wenti_data.point_no(2), t_coord, True)
 End If
 ' poi(wenti_data.point_no(2)).data(0).is_change
 '  Call change_circle_(wenti_data.point_no(2), pointapi0)
End If

End Function

Private Function change_picture_6(wenti_data As wentitype, change_element As condition_type) As Boolean
'□□=!_~
Dim r!
change_picture_6 = True
If wenti_data.point_no(8) > 0 Then
 If wenti_data.point_no(0) = yidian_no Then
  'r! = Sqr((move_x& - m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.X) ^ 2 + _
      (move_y& - m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.Y) ^ 2)
     t_coord.X = mouse_move_coord.X + _
           wenti_data.point_no(17) '
     t_coord.Y = mouse_move_coord.Y + _
           wenti_data.point_no(18) '
             Call set_point_coordinate(wenti_data.point_no(1), t_coord, True) '
 ElseIf wenti_data.point_no(1) = yidian_no Then '
  r! = sqr((move_coord.X - m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X) ^ 2 + _
      (move_coord.Y - m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y) ^ 2)
    Call C_display_wenti.set_m_point_no(wenti_data.wenti_no, _
    m_Circ(wenti_data.point_no(8)).data(0).data0.radii * _
          (move_coord.X - m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X) / r!, 17, False)
    Call C_display_wenti.set_m_point_no(wenti_data.wenti_no, _
          m_Circ(wenti_data.point_no(8)).data(0).data0.radii * _
          (move_coord.Y - m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y) / r!, 18, False)
    t_coord.X = m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X _
      + wenti_data.point_no(17)
    t_coord.Y = m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y _
     + wenti_data.point_no(18)
       Call set_point_coordinate(wenti_data.point_no(1), t_coord, True)
End If
End If

End Function

Private Function change_picture_2(wenti_data As wentitype) As Boolean
'-2 作⊙□[down\\(_)]和⊙□[down\\(_)]的公切线□□
'-60 作⊙□□□和⊙□□□的公切线□□
'-59 作⊙□□□和⊙□[down\\(_)]的公切线□□
Dim i%, n%
Dim ty As Boolean
change_picture_2 = True
'If wenti_data.poi(1) > yidian_no Then
'If m_Circ(wenti_data.circ(1)).data(0).is_change Or
'     m_Circ(wenti_data.circ(2)).data(0).is_change Then   '圆移动
'Call change_tangent_line_for_two_circle(wenti_data.line_no(1))
'******
'End If
End Function

Private Function change_picture_1_30_31_58(wenti_data As wentitype, change_element As condition_type) As Boolean
'□□＝□□
Dim inter_type As Integer
Dim out_coord(1) As POINTAPI
  If (change_element_ty = point_ And change_element_no = wenti_data.poi(3)) Or _
       (change_element_ty = point_ And change_element_no = wenti_data.poi(4)) Then '标准线段变化
       m_Circ(wenti_data.circ(1)).data(0).data0.radii = _
        distance_of_two_POINTAPI(m_poi(wenti_data.poi(3)).data(0).data0.coordinate, _
          m_poi(wenti_data.poi(4)).data(0).data0.coordinate)
              m_Circ(wenti_data.circ(1)).data(0).is_change = True
  End If
  If wenti_data.no = -31 Then '
   If (change_element_ty = circle_ And change_element_no = (wenti_data.circ(1)) Or _
        change_element_ty = line_ And change - element_no = wenti_data.line_no(1)) Then
        'If m_Circ(wenti_data.circ(1)).data(0).is_change Then
        '    Call change_m_circle(wenti_data.circ(1), depend_condition(wenti_cond_, wenti_data.wenti_no))
        'End If
    Call inter_point_line_circle(wenti_data.line_no(1), wenti_data.circ(1), _
             m_poi(wenti_data.poi(1)).data(0).data0.coordinate, wenti_data.poi(1), _
               False, False, wenti_data.inter_set_point_type)
        'm_poi(wenti_data.poi(1)).data(0).is_change = True
           Call change_m_point(wenti_data.poi(1), depend_condition(wenti_cond_, wenti_data.wenti_no))
    End If
  Else '30,58
    inter_type = inter_point_circle_circle_(m_Circ(wenti_data.circ(1)).data(0).data0, m_Circ(wenti_data.circ(2)).data(0).data0, _
               out_coord(0), 0, out_coord(1), 0, 0, 0, True)
            If inter_type = 2 Then
            If wenti_data.inter_set_point_type = new_point_on_circle_circle12 Then
              m_poi(wenti_data.poi(1)).data(0).data0.coordinate = out_coord(0)
            ElseIf wenti_data.inter_set_point_type = new_point_on_circle_circle21 Then
              m_poi(wenti_data.poi(1)).data(0).data0.coordinate = out_coord(1)
            End If
            Else
            m_poi(wenti_data.poi(1)).data(0).data0.coordinate = out_coord(0)
            End If
             'm_poi(wenti_data.poi(1)).data(0).is_change = True
             Call change_m_point(wenti_data.poi(1), depend_condition(wenti_cond_, wenti_data.wenti_no))
  End If
End Function
Private Sub change_picture0(ByVal num As Integer)
'Case 0
Dim m%, i%
Dim A!
Dim r!
Dim b!
Dim p!
Dim q!
Dim s!
Dim t!
Dim g!
If yidian_no > 0 Then
   m% = 0
For i% = 0 To 5
  If addition_condition_statue Then
    If m_poi(yidian_no).data(0).degree = 1 Then
     '附加的等式条件
     If C_display_wenti.m_point_no(num, i%) = yidian_no Then
       Call C_display_wenti.set_m_point_no(num, addition_condition(last_addition_condition), 5 + i%, False)
        m% = addition_condition(last_addition_condition)
         mouse_move_coord = m_poi(yidian_no).data(0).data0.coordinate
          ' move_y& = poi(yidian_no).data(0).data0.coordinate.Y
            addition_condition_statue = False
             GoTo change_picture_mark01
   End If
   End If
 Else
  If C_display_wenti.m_point_no(num, 5 + i%) > 0 Then
     m% = C_display_wenti.m_point_no(num, 5 + i%)
     GoTo change_picture_mark01
  End If
End If
Next i%
change_picture_mark01:
If C_display_wenti.m_no(num) = -1 Then
'm%记录调整点所在的语句
  If C_display_wenti.m_point_no(num, 4) > yidian_no Then
  A! = C_display_wenti.m_point_no(num, 9) / 1000
   b! = C_display_wenti.m_point_no(num, 10) / 1000
  If m_Circ(C_display_wenti.m_point_no(num, 6)).data(0).is_change Then
  If m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.center > 0 Then
  m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.radii = _
  sqr(CSng(m_poi(C_display_wenti.m_point_no(num, 6)).data(0).data0.coordinate.X _
   - m_poi(C_display_wenti.m_point_no(num, 7)).data(0).data0.coordinate.X) ^ 2 + _
     CSng(m_poi(C_display_wenti.m_point_no(num, 6)).data(0).data0.coordinate.Y - _
         m_poi(C_display_wenti.m_point_no(num, 7)).data(0).data0.coordinate.Y) ^ 2)
  Else
  m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.radii = _
  sqr(CSng(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.c_coord.X _
   - m_poi(C_display_wenti.m_point_no(num, 7)).data(0).data0.coordinate.X) ^ 2 + _
     CSng(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.c_coord.Y - _
         m_poi(C_display_wenti.m_point_no(num, 7)).data(0).data0.coordinate.Y) ^ 2)
  End If
       mouse_move_coord = m_poi(C_display_wenti.m_point_no(num, 4)).data(0).data0.coordinate '
         'move_y& = poi(C_display_wenti.m_point_no(num,4)).data(0).data0.coordinate.Y
   End If
  ElseIf C_display_wenti.m_point_no(num, 4) = yidian_no And _
    m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(1) = _
     C_display_wenti.m_point_no(num, 4) Then
   r! = sqr(((move_coord.X - m_poi(C_display_wenti.m_point_no(num, 5)).data(0).data0.coordinate.X)) ^ 2 + _
     (move_coord.Y - m_poi(C_display_wenti.m_point_no(num, 5)).data(0).data0.coordinate.Y) ^ 2)
  A! = (move_coord.X - m_poi(C_display_wenti.m_point_no(num, 5)).data(0).data0.coordinate.X) / r!
   b! = (move_coord.Y - m_poi(C_display_wenti.m_point_no(num, 5)).data(0).data0.coordinate.Y) / r!
     Call C_display_wenti.set_m_point_no(num, CInt(A! * 1000), 9, False)
     Call C_display_wenti.set_m_point_no(num, CInt(b! * 1000), 10, False)
Else
End If
If m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(1) = _
 C_display_wenti.m_point_no(num, 4) Then
   t_coord.X = _
     m_poi(C_display_wenti.m_point_no(num, 5)).data(0).data0.coordinate.X + _
       CInt(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.radii * A!)
   t_coord.Y = _
     m_poi(C_display_wenti.m_point_no(num, 5)).data(0).data0.coordinate.Y + _
       CInt(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.radii * b!)
  Call set_point_coordinate(C_display_wenti.m_point_no(num, 4), t_coord, True)
Else '　是圆心
 'If Abs(move_x% - m_poi(c_display_wenti.m_point_no(num,4)).data(0).data0.coordinate.X) < 8 And _
  Abs(move_y% - m_poi(c_display_wenti.m_point_no(num,4)).data(0).data0.coordinate.Y) < 8 Then
r! = (CSng(m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(1)).data(0).data0.coordinate.X - _
 m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(2)).data(0).data0.coordinate.X)) ^ 2 + _
  (m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(1)).data(0).data0.coordinate.Y - _
   m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(2)).data(0).data0.coordinate.Y) ^ 2
p! = (m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(1)).data(0).data0.coordinate.X + _
 m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(2)).data(0).data0.coordinate.X) / 2 - _
  mouse_move_coord.X
q! = (m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(1)).data(0).data0.coordinate.Y + _
 m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(2)).data(0).data0.coordinate.Y) / 2 - _
 mouse_move_coord.Y
s! = m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(2)).data(0).data0.coordinate.X - _
 m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(1)).data(0).data0.coordinate.X
t! = m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(2)).data(0).data0.coordinate.Y - _
 m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(1)).data(0).data0.coordinate.Y
 g! = (p! * s! + q! * t!) / r!
  t_coord.X = _
    mouse_move_coord.X + _
         (m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(2)).data(0).data0.coordinate.X - _
        m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(1)).data(0).data0.coordinate.X) * g!
  t_coord.Y = _
    mouse_move_coord.Y + _
       (m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(2)).data(0).data0.coordinate.Y - _
        m_poi(m_Circ(C_display_wenti.m_point_no(num, 8)).data(0).data0.in_point(1)).data(0).data0.coordinate.Y) * g!
  Call set_point_coordinate(C_display_wenti.m_point_no(num, 4), t_coord, True)
End If
End If
End If
End Sub

Private Function change_picture1(wenti_data As wentitype, change_element As condition_type, t_cp%, ByVal num As Integer) As Boolean
'直线□□上任取一点□
Dim tp%, tl%
Dim A!
change_picture1 = True
'If wenti_data.point_no(29) >= change_p% And wenti_data.point_no(28) <= change_p% Then
'If m_poi(wenti_data.poi(1)).data(0).degree > 0 Then
'If m_lin(wenti_data.line_no(1)).data(0).is_change Then
  tp% = wenti_data.poi(1)
  tl% = wenti_data.line_no(1)
  m_poi(tp%).data(0).data0.coordinate = add_POINTAPI(m_poi(m_lin(tl%).data(0).data0.poi(1)).data(0).data0.coordinate, _
            time_POINTAPI_by_number(minus_POINTAPI(m_poi(m_lin(tl%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                         m_poi(m_lin(tl%).data(0).data0.poi(1)).data(0).data0.coordinate), m_poi(tp%).data(0).parent.ratio))
  't_coord.X = _
     m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X + _
 wenti_data.point_no(3) * (m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.X - _
  m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X) / 1000
  't_coord.Y = m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y + _
 wenti_data.point_no(3) * (m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.Y - _
  m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y) / 1000
  'Call set_point_coordinate(wenti_data.poi(1), t_coord, True)
 'm_poi(tp%).data(0).is_change = True
 Call change_m_point(tp%)
'ElseIf m_poi(wenti_data.poi(1)).data(0).is_change Then
'    If read_line1(m_poi(wenti_data.point_no(0)).data(0).data0.coordinate, _
     m_poi(wenti_data.point_no(1)).data(0).data0.coordinate, _
      m_poi(change_p%).data(0).data0.coordinate, t_coord, wenti_data.poi(1), A!, 5, True) Then
'       Call C_display_wenti.set_m_point_no(num, Int(A! * 1000), 3, False)
        'Call set_point_coordinate(C_display_wenti.m_point_no(2), t_coord, True)
 '    End If
'End If
'End If
'If wenti_data.poi(1) < wenti_data.point_no(29) Then
' t_cp% = wenti_data.poi(1)
'End If
'End If
End Function

Private Function change_picture2_3(wenti_data As wentitype) As Boolean
'2 □□∥□□
'3 □□⊥□□
Dim t_cond As condition_data_type
Dim p_coord(1) As POINTAPI
Dim t_coord As POINTAPI
'Dim paral_or_verti As Integer
'If wenti_data.no = 2 Then
' paral_or_verti = paral_
'Else
' paral_or_verti = verti_
'End If
change_picture2_3 = True
 If m_poi(wenti_data.poi(1)).data(0).is_change Then '平行线的后端点变化
   Call distance_point_to_line(m_poi(wenti_data.poi(1)).data(0).data0.coordinate, _
                                m_poi(wenti_data.poi(2)).data(0).data0.coordinate, _
                                 wenti_data.no, _
                                  m_poi(m_lin(wenti_data.line_no(1)).data(0).data0.poi(1)).data(0).data0.coordinate, _
                                   m_poi(m_lin(wenti_data.line_no(1)).data(0).data0.poi(0)).data(0).data0.coordinate, _
                                    0, m_poi(wenti_data.poi(1)).data(0).data0.coordinate) '计算后端点的坐标，投影到平行线上
'设置平行线与原线段的长度比
           If wenti_data.no = paral_ Then
              t_coord = minus_POINTAPI(m_poi(wenti_data.poi(3)).data(0).data0.coordinate, _
                                m_poi(wenti_data.poi(4)).data(0).data0.coordinate)
           Else 'If w_n = 3 Then
             t_coord = verti_POINTAPI(minus_POINTAPI(m_poi(wenti_data.poi(3)).data(0).data0.coordinate, _
                                m_poi(wenti_data.poi(4)).data(0).data0.coordinate))
          End If
          If Abs(t_coord.X) > 5 Then
           m_poi(wenti_data.poi(1)).data(0).parent.ratio = (m_poi(wenti_data.poi(1)).data(0).data0.coordinate.X - _
                                   m_poi(wenti_data.poi(2)).data(0).data0.coordinate.X) / t_coord.X
          Else
           m_poi(wenti_data.poi(1)).data(0).parent.ratio = (m_poi(wenti_data.poi(1)).data(0).data0.coordinate.Y - _
                                   m_poi(wenti_data.poi(2)).data(0).data0.coordinate.Y) / t_coord.Y
         End If
 '*******************************************************************************************************************************
            Call arrange_move_point_on_line(wenti_data.poi(1), wenti_data.line_no(2)) '整理平行线上的点的排列
           'm_lin(wenti_data.line_no(2)).data(0).is_change = True
           Call change_m_line(wenti_data.line_no(2))
 Else '原直线或平行线起点变化，计算平行线的末端坐标
   If wenti_data.no = 2 Then
    m_poi(wenti_data.poi(1)).data(0).data0.coordinate = _
           add_POINTAPI(m_poi(wenti_data.poi(2)).data(0).data0.coordinate, _
                  time_POINTAPI_by_number(minus_POINTAPI( _
                  m_poi(m_lin(wenti_data.line_no(1)).data(0).data0.poi(1)).data(0).data0.coordinate, _
                   m_poi(m_lin(wenti_data.line_no(1)).data(0).data0.poi(0)).data(0).data0.coordinate), _
                     m_poi(wenti_data.poi(1)).data(0).parent.ratio))  '计算平行线的端点
  Else
    m_poi(wenti_data.poi(1)).data(0).data0.coordinate = _
     add_POINTAPI(m_poi(wenti_data.poi(2)).data(0).data0.coordinate, _
                    time_POINTAPI_by_number(verti_POINTAPI(minus_POINTAPI( _
                  m_poi(m_lin(wenti_data.line_no(1)).data(0).data0.poi(1)).data(0).data0.coordinate, _
                   m_poi(m_lin(wenti_data.line_no(1)).data(0).data0.poi(0)).data(0).data0.coordinate)), _
                    m_poi(wenti_data.poi(1)).data(0).parent.ratio))
  End If
           '  m_lin(wenti_data.line_no(2)).data(0).is_change = True
          '   m_poi(wenti_data.poi(1)).data(0).is_change = True
          '   Call change_m_point(wenti_data.poi(1))
          '   Call arrange_move_point_on_line(wenti_data.poi(1), wenti_data.line_no(1))
              'm_lin(wenti_data.line_no(2)).data(0).is_change = True
              Call change_m_line(wenti_data.line_no(2))
End If
End Function
Private Function change_picture4(wenti_data As wentitype, change_element As condition_type, t_cp%) As Boolean
'4在□□的垂直平分线上任取一点□
Dim m%
Dim A!
Dim r&
Dim mid_coord As POINTAPI
Dim r1&
If wenti_data.no_ = -54 Or wenti_data.no_ = -53 Then
   wenti_data.no = wenti_data.no_
  change_picture4 = change_picture_54_23_22(wenti_data)
Else
change_picture4 = True
 If change_element.ty = point_ And change_element.no = wenti_data.poi(1) Then
 Call distance_point_to_line(m_poi(wenti_data.poi(1)).data(0).data0.coordinate, _
                 m_poi(wenti_data.poi(2)).data(0).data0.coordinate, verti_, _
                  m_poi(wenti_data.poi(3)).data(0).data0.coordinate, _
                   m_poi(wenti_data.poi(4)).data(0).data0.coordinate, _
                    0, m_poi(wenti_data.poi(1)).data(0).data0.coordinate)
  t_coord = verti_POINTAPI(minus_POINTAPI(m_poi(wenti_data.poi(4)).data(0).data0.coordinate, _
                                            m_poi(wenti_data.poi(3)).data(0).data0.coordinate)) '4-3
   'm_poi(wenti_data.poi(1)).data(0).data0.coordinate = _
              add_POINTAPI(m_poi(wenti_data.poi(2)).data(0).data0.coordinate, _
                time_POINTAPI_by_number(t_coord, m_poi(wenti_data.poi(1)).data(0).parent.ratio))
           Call arrange_move_point_on_line(wenti_data.poi(1), wenti_data.line_no(2)) '整理平行线上的点的排列
        If Abs(t_coord.X) > 5 Then '
          m_poi(wenti_data.poi(1)).data(0).parent.ratio = (m_poi(wenti_data.poi(1)).data(0).data0.coordinate.X - _
                           m_poi(wenti_data.poi(2)).data(0).data0.coordinate.X) / t_coord.X '1-2/4-3
        Else
          m_poi(wenti_data.poi(1)).data(0).parent.ratio = (m_poi(wenti_data.poi(1)).data(0).data0.coordinate.Y - _
                             m_poi(wenti_data.poi(2)).data(0).data0.coordinate.Y) / t_coord.Y '1-2/4-3
        End If
           Call arrange_move_point_on_line(wenti_data.poi(4), wenti_data.line_no(2))
          ' m_lin(wenti_data.line_no(2)).data(0).is_change = True
           Call change_m_line(wenti_data.line_no(2))

'*****************************************************************************************************

ElseIf (change_element.ty = point_ And change_element.no = wenti_data.poi(3)) Or _
     (change_element.ty = point_ And change_element.ty = point_ And change_element.no = wenti_data.poi(4)) Then '
m_poi(wenti_data.poi(2)).data(0).data0.coordinate = _
  mid_POINTAPI(m_poi(wenti_data.poi(3)).data(0).data0.coordinate, _
    m_poi(wenti_data.poi(4)).data(0).data0.coordinate) '求中点
    '************************************************************
    'm_poi(wenti_data.poi(2)).data(0).is_change = True
    Call change_m_point(wenti_data.poi(2))
    '**********************************************************************
 m_poi(wenti_data.poi(1)).data(0).data0.coordinate = _
         add_POINTAPI(m_poi(wenti_data.poi(2)).data(0).data0.coordinate, _
          time_POINTAPI_by_number(verti_POINTAPI(minus_POINTAPI(m_poi(wenti_data.poi(4)).data(0).data0.coordinate, _
          m_poi(wenti_data.poi(3)).data(0).data0.coordinate)), m_poi(wenti_data.poi(1)).data(0).parent.ratio))
   ' Call arrange_move_point_on_line(wenti_data.poi(1), wenti_data.line_no(1))
    'm_poi(wenti_data.poi(1)).data(0).is_change = True
   Call change_m_point(wenti_data.poi(1))
'**********************************************************************************************************
  'm_poi(wenti_data.poi(1)).data(0).is_change = True
   Call change_m_point(wenti_data.poi(1))
End If
End If
End Function

Private Function change_picture5(wenti_data As wentitype, _
                                  change_element As condition_type, t_cp%) As Boolean
'线段□□的中点□,移动线段端点change_p%
change_picture5 = True
 If (change_element.ty = point_ And change_element.no = wenti_data.poi(2)) Or _
      (change_element.ty = point_ And change_element.no = wenti_data.poi(3)) Then
      '如果线段有一个端点移动，那么重新计算线段重点
       m_poi(wenti_data.poi(1)).data(0).data0.coordinate = _
         mid_POINTAPI(m_poi(wenti_data.poi(2)).data(0).data0.coordinate, _
           m_poi(wenti_data.poi(3)).data(0).data0.coordinate) '计算中点坐标
  'm_poi(wenti_data.poi(1)).data(0).is_change = True '设置中点移动属性真
  Call change_m_point(wenti_data.poi(1))   '移动线段中点 ，中点不能在此直线上移动
End If
End Function

Private Function change_picture6(wenti_data As wentitype) As Boolean
'6 □是线段□□上分比为!_~的分点
Dim v As Single
change_picture6 = True
'If (change_element.ty = point_ And change_element.no = wenti_data.poi(2)) Or _
       change_element.ty = point_ And change_element.no = wenti_data.poi(3) Then
'Call C_display_wenti.Get_wenti(wenti_data.wenti_no)
If m_poi(wenti_data.point_no(1)).data(0).is_change Or _
   m_poi(wenti_data.point_no(2)).data(0).is_change Then '线段至少有一个端点移动
     v = value_for_draw(wenti_data.point_no(3)) '读取比值
 m_poi(wenti_data.poi(1)).data(0).data0.coordinate = _
   add_POINTAPI(m_poi(wenti_data.point_no(1)).data(0).data0.coordinate, _
        divide_POINTAPI_by_number(minus_POINTAPI(m_poi(wenti_data.point_no(2)).data(0).data0.coordinate, _
         m_poi(wenti_data.point_no(1)).data(0).data0.coordinate), v)) '计算分点坐标
     '计算定比分点的坐标
      'm_poi(wenti_data.poi(1)).data(0).is_change = True
      Call change_m_point(wenti_data.poi(1)) '移动定比分点
'End If
End If
End Function

Private Function change_picture7_61(wenti_data As wentitype) As Boolean
'⊙□[down\\(_)]上任取一点□
Dim r!
If wenti_data.point_no(29) >= yidian_no And wenti_data.point_no(28) <= yidian_no Then
change_picture7_61 = True
If m_Circ(wenti_data.circ(1)).data(0).is_change Or _
     m_poi(wenti_data.poi(1)).data(0).is_change Then
      Call put_point_to_circle(m_poi(wenti_data.poi(1)).data(0).data0.coordinate, _
        m_Circ(wenti_data.circ(1)).data(0).data0, _
         t_coord, wenti_data.poi(1))
End If
End If
End Function

Private Function change_picture9(wenti_data As wentitype, change_element As condition_type, c_p%) As Boolean
'直线□□和直线□□交于点□
Dim t_cond As condition_data_type
change_picture9 = True
 If (change_element.ty = line_ And change_element.no = wenti_data.line_no(1)) Or _
      (change_element.ty = line_ And change_element.no = wenti_data.line_no(2)) Then
     '计算两直线的交点坐标
     Call calculate_line_line_intersect_point( _
      m_poi(m_lin(wenti_data.line_no(1)).data(0).data0.poi(0)).data(0).data0.coordinate, _
      m_poi(m_lin(wenti_data.line_no(1)).data(0).data0.poi(1)).data(0).data0.coordinate, _
      m_poi(m_lin(wenti_data.line_no(2)).data(0).data0.poi(0)).data(0).data0.coordinate, _
      m_poi(m_lin(wenti_data.line_no(2)).data(0).data0.poi(1)).data(0).data0.coordinate, _
      m_poi(wenti_data.poi(1)).data(0).data0.coordinate, 0)
       'm_poi(wenti_data.poi(1)).data(0).is_change = True
       Call change_m_point(wenti_data.poi(1))
     End If
End Function

Private Function change_picture10_16(wenti_data As wentitype, change_element As condition_type) As Boolean
'10 过□平行□□的直线交⊙□[down\\(_)]于□
'16 过□垂直□□的直线交⊙□[down\\(_)]于□
'-68 过□垂直□□的直线交⊙□□□于□
'-62 过□平行□□的直线交⊙□□□于□
'-53□□垂直平分线交⊙□[down\\(_)]于□
'-25
Dim ty As Integer
Dim tcoord As POINTAPI
If wenti_data.poi(1) > yidian_no Then
change_picture10_16 = True
If wenti_data.no = 10 Or wenti_data.no = -62 Then
 ty = paral_
ElseIf wenti_data.no = 16 Or wenti_data.no = -68 Then
 ty = verti_
ElseIf wenti_data.no = -53 Or wenti_data.no = -25 Then
 m_poi(wenti_data.poi(2)).data(0).data0.coordinate = mid_POINTAPI(m_poi(wenti_data.point_no(0)).data(0).data0.coordinate, _
                  m_poi(wenti_data.point_no(1)).data(0).data0.coordinate)
 ty = verti_
 'm_poi(wenti_data.poi(2)).data(0).is_change = True
 Call change_m_point(wenti_data.poi(2))
End If
'If m_poi(wenti_data.point_no(0)).data(0).is_change Or _
    m_lin(wenti_data.line_no(1)).data(0).is_change Then
    Call change_paral_or_verti_line(wenti_data.line_no(1), wenti_data.poi(2), _
          wenti_data.line_no(2), ty)
           'm_lin(wenti_data.line_no(2)).data(0).is_change = True
            Call change_m_line(wenti_data.line_no(2))
'End If
Call inter_point_line_circle(wenti_data.line_no(2), wenti_data.circ(1), _
     m_poi(wenti_data.poi(1)).data(0).data0.coordinate, wenti_data.poi(1), _
      False, False, wenti_data.inter_set_point_type)
     'm_poi(wenti_data.poi(1)).data(0).is_change = True
     Call change_m_point(wenti_data.poi(1))
End If
End Function
Private Function change_picture11(wenti_data As wentitype) As Boolean
'11□是直线□□与⊙□[down\\(_)]的一个交点
'-63
Dim X&, Y&, x1&, y1&
change_picture11 = True
If m_lin(wenti_data.line_no(1)).data(0).is_change Or _
    m_Circ(wenti_data.circ(1)).data(0).is_change = 255 Then
Call inter_point_line_circle(wenti_data.line_no(1), wenti_data.circ(1), _
             m_poi(wenti_data.poi(1)).data(0).data0.coordinate, wenti_data.poi(1), _
               False, False, wenti_data.inter_set_point_type)
 'm_poi(wenti_data.poi(1)).data(0).is_change = True
  Call change_m_point(wenti_data.poi(1))
End If
End Function

Private Function change_picture12(wenti_data As wentitype) As Boolean
'⊙□[down\\(_)]和⊙□[down\\(_)]相切于点□
'-65 ⊙□□□和⊙□□□相切于点□
'-64 ⊙□□□和⊙□[down\\(_)]相切于点□
Dim r!, d!
change_picture12 = True
If wenti_data.poi(1) > yidian_no Then
If m_Circ(wenti_data.circ(1)).data(0).is_change Or m_Circ(wenti_data.circ(2)).data(0).is_change Then
   d! = sqr((m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.X - _
             m_Circ(wenti_data.circ(2)).data(0).data0.c_coord.X) ^ 2 + _
            (m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.Y - _
             m_Circ(wenti_data.circ(2)).data(0).data0.c_coord.Y) ^ 2) ' 圆心距
If m_Circ(wenti_data.circ(2)).data(0).data0.center < wenti_data.poi(1) Then
 t_coord = _
   add_POINTAPI(m_Circ(wenti_data.circ(1)).data(0).data0.c_coord, _
     time_POINTAPI_by_number(minus_POINTAPI(m_Circ(wenti_data.circ(2)).data(0).data0.c_coord, _
            m_Circ(wenti_data.circ(1)).data(0).data0.c_coord), _
                  m_Circ(wenti_data.circ(1)).data(0).data0.radii / d!))
 Call set_point_coordinate(wenti_data.poi(1), t_coord, True) '改变切点
ElseIf wenti_data.point_no(9) = 3 Then
'   r! = sqr((m_poi(wenti_data.poi(1)).data(0).data0.coordinate.X - _
          m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.X) ^ 2 + _
          (m_poi(wenti_data.poi(1)).data(0).data0.coordinate.Y - _
            m_Circ(wenti_data(1)).data(0).data0.c_coord.Y) ^ 2) '
't_coord.X = _
 m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.X + _
 (m_poi(wenti_data.poi(1)).data(0).data0.coordinate.X - _
   m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.X) * _
     m_Circ(wenti_dara.circ(1)).data(0).data0.radii / r!
't_coord.Y = _
 m_Circ(wenti_data.m_point_no(num,12)).data(0).data0.c_coord.Y + _
 (m_poi(wenti_data.m_point_no(num,4)).data(0).data0.coordinate.Y - _
   m_Circ(wenti_data.m_point_no(num,12)).data(0).data0.c_coord.Y) * _
    m_Circ(wenti_data.m_point_no(num,12)).data(0).data0.radii / r!
    Call set_point_coordinate(wenti_data.point_no(4), t_coord, True)
      'poi(wenti_data.m_point_no(num,4)).data(0).is_change
       '  Call change_circle_(wenti_data.m_point_no(num,4), pointapi0)
'r! = sqr((m_Circ(wenti_data.m_point_no(num,13)).data(0).data0.c_coord.X - _
          m_Circ(wenti_data.m_point_no(num,12)).data(0).data0.c_coord.X) ^ 2 + _
        (m_Circ(wenti_data.m_point_no(num,13)).data(0).data0.c_coord.Y - _
          m_Circ(wenti_data.m_point_no(num,12)).data(0).data0.c_coord.Y) ^ 2)
    t_coord = _
     add_POINTAPI(m_Circ(wenti_data.circ(1)).data(0).data0.c_coord, _
          time_POINTAPI_by_number(minus_POINTAPI(m_poi(wenti_data.poi(1)).data(0).data0.coordinate, _
        m_Circ(wenti_data.circ(1)).data(0).data0.c_coord), r! / _
            m_Circ(wenti_data.circ(1)).data(0).data0.radii))
    't_coord.Y = _
     m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.Y + _
      (m_poi(wenti_data.poi(1)).data(0).data0.coordinate.Y - _
        m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.Y) * r! / _
         m_Circ(wenti_data.circ(1)).data(0).data0.radii
          Call set_point_coordinate(m_Circ(wenti_data.circ(2)).data(0).data0.center, _
             t_coord, True)
     'poi(m_circ(wenti_data.m_point_no(num,13)).data(0).data0.center).data(0).is_change
       '  Call change_circle_(m_circ(wenti_data.m_point_no(num,13)).data(0).data0.center, pointapi0)
'ElseIf wenti_data.m_point_no(num,9) = 11 Then
' Call set_tangent_circles(m_Circ(wenti_data.m_point_no(num,12)).data(0).data0, _
        m_poi(m_Circ(wenti_data.m_point_no(num,13)).data(0).data0.in_point(1)).data(0).data0.coordinate, _
         m_poi(m_Circ(wenti_data.m_point_no(num,13)).data(0).data0.in_point(2)).data(0).data0.coordinate, _
          m_Circ(wenti_data.m_point_no(num,13)).data(0).data0.radii, _
           0, m_Circ(wenti_data.m_point_no(num,13)).data(0).data0.c_coord, _
            pointapi0, t_coord1, wenti_data.m_point_no(num,4), _
              pointapi0, 0, 1, True)
        'poi(wenti_data.m_point_no(num,4)).data(0).is_change
        '  Call change_circle_(wenti_data.m_point_no(num,4), pointapi0)
'ElseIf wenti_data.m_point_no(num,9) = 12 Then
'  Call set_tangent_circles(m_Circ(wenti_data.m_point_no(num,12)).data(0).data0, _
        m_poi(m_Circ(wenti_data.m_point_no(num,13)).data(0).data0.in_point(1)).data(0).data0.coordinate, _
         m_poi(m_Circ(wenti_data.m_point_no(num,13)).data(0).data0.in_point(2)).data(0).data0.coordinate, _
          0, m_Circ(wenti_data.m_point_no(num,13)).data(0).data0.radii, _
           pointapi0, m_Circ(wenti_data.m_point_no(num,13)).data(0).data0.c_coord, _
            t_coord1, 0, t_coord2, wenti_data.m_point_no(num,4), _
               2, True)
End If
t_coord.X = _
    m_poi(wenti_data.poi(1)).data(0).data0.coordinate.X + _
     (m_Circ(wenti_data.circ(2)).data(0).data0.c_coord.Y - _
       m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.Y)
t_coord.Y = _
     m_poi(wenti_data.poi(1)).data(0).data0.coordinate.Y + _
     (m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.X - _
       m_Circ(wenti_data.circ(2)).data(0).data0.c_coord.X)
      Call set_point_coordinate(wenti_data.poi(2), t_coord, True)
t_coord.X = _
     m_poi(wenti_data.poi(1)).data(0).data0.coordinate.X + _
     (m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.Y - _
       m_Circ(wenti_data.circ(2)).data(0).data0.c_coord.Y)
t_coord.Y = _
     m_poi(wenti_data.poi(1)).data(0).data0.coordinate.Y + _
     (m_Circ(wenti_data.circ(2)).data(0).data0.c_coord.X - _
       m_Circ(wenti_data.circ(1)).data(0).data0.c_coord.X)
      Call set_point_coordinate(wenti_data.poi(3), t_coord, True)
End If
End If
End Function

Private Function change_picture13(wenti_data As wentitype) As Boolean
'13 □是⊙□[down\\(_)]和⊙□[down\\(_)]一个交点
'-67 □是⊙□□□和⊙□□□一个交点
'-66 □是⊙□□□和⊙□[down\\(_)]一个交点
Dim X&, Y&
Dim i%
Dim out_coord(1) As POINTAPI
change_picture13 = True
'If wenti_data.poi(1) > yidian_no Then
'******************************************************************
If m_Circ(wenti_data.circ(1)).data(0).is_change Or _
    m_Circ(wenti_data.circ(2)).data(0).is_change Then
 Call inter_point_circle_circle_( _
     m_Circ(wenti_data.circ(1)).data(0).data0, m_Circ(wenti_data.circ(2)).data(0).data0, _
         out_coord(0), 0, out_coord(1), 0, 0, 0, False)
     If wenti_data.inter_set_point_type = new_point_on_circle_circle12 Then
              m_poi(wenti_data.poi(1)).data(0).data0.coordinate = out_coord(0)
     ElseIf wenti_data.inter_set_point_type = new_point_on_circle_circle21 Then
              m_poi(wenti_data.poi(1)).data(0).data0.coordinate = out_coord(1)
     End If
   'm_poi(wenti_data.poi(1)).data(0).is_change = True
   Call change_m_point(wenti_data.poi(1))
'End If
End If

End Function

Private Function change_picture14(wenti_data As wentitype) As Boolean
'过□作直线□□的垂线垂足为□
change_picture14 = True
If wenti_data.poi(1) > yidian_no Then
If m_poi(wenti_data.point_no(0)).data(0).is_change Or _
    m_poi(wenti_data.point_no(1)).data(0).is_change Or _
     m_poi(wenti_data.point_no(2)).data(0).is_change Then
  Call orthofoot1( _
   m_poi(wenti_data.point_no(0)).data(0).data0.coordinate, _
    m_poi(wenti_data.point_no(1)).data(0).data0.coordinate, _
      m_poi(wenti_data.point_no(2)).data(0).data0.coordinate, _
        t_coord, wenti_data.point_no(3), True)
End If
End If
End Function

Private Function change_picture15(wenti_data As wentitype) As Boolean '2003
'15 以□□为直径作⊙□[down\\(_)]
change_picture15 = True
If wenti_data.poi(1) > yidian_no Then
 If m_poi(wenti_data.point_no(0)).data(0).is_change Or _
               m_poi(wenti_data.point_no(1)).data(0).is_change Then
    t_coord.X = ( _
    m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.X + _
    m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.X) / 2
    t_coord.Y = ( _
    m_poi(wenti_data.point_no(0)).data(0).data0.coordinate.Y + _
    m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.Y) / 2
    Call set_point_coordinate(wenti_data.point_no(2), t_coord, True)
    Call C_display_picture.set_circle_center(wenti_data.point_no(0), _
            wenti_data.point_no(2))
 End If
End If

End Function

Private Function change_picture18(wenti_data As wentitype) As Boolean
'18 □是△□□□的重心
change_picture18 = True
 If m_poi(wenti_data.point_no(1)).data(0).is_change Or _
     m_poi(wenti_data.point_no(2)).data(0).is_change Or _
      m_poi(wenti_data.point_no(3)).data(0).is_change Then
Call centroid1(wenti_data.point_no(1), _
   wenti_data.point_no(2), _
      wenti_data.point_no(3), _
         wenti_data.point_no(0), _
            wenti_data.point_no(4), _
               wenti_data.point_no(5), _
                 wenti_data.point_no(6), True)
 End If
 End Function

Private Function change_picture19(wenti_data As wentitype) As Boolean
'19 □是△□□□的垂心
change_picture19 = True
If m_poi(wenti_data.point_no(1)).data(0).is_change Or _
     m_poi(wenti_data.point_no(2)).data(0).is_change Or _
       m_poi(wenti_data.point_no(3)).data(0).is_change Then
Call redraw_three_point_circle(m_Circ(wenti_data.point_no(10)).data(0).data0, True)
     'poi(wenti_data.point_no(0)).data(0).is_change
     '    Call change_circle_(wenti_data.point_no(0), pointapi0)
  End If
End Function
Private Function change_picture20(wenti_data As wentitype) As Boolean
'20 □是△□□□的垂心
'Case 20 '垂心
change_picture20 = True
If m_poi(wenti_data.point_no(1)).data(0).is_change Or _
    m_poi(wenti_data.point_no(2)).data(0).is_change Or _
      m_poi(wenti_data.point_no(3)).data(0).is_change Then
 Call orthocenter(wenti_data.point_no(1), _
   wenti_data.point_no(2), _
     wenti_data.point_no(3), _
       wenti_data.point_no(0), _
        wenti_data.point_no(4), _
          wenti_data.point_no(5), _
            wenti_data.point_no(6), 0, True)
End If
End Function

Private Function change_picture21(wenti_data As wentitype) As Boolean
'21 □是△□□□的内切圆的圆心
change_picture21 = True
 If m_poi(wenti_data.point_no(1)).data(0).is_change Or _
     m_poi(wenti_data.point_no(2)).data(0).is_change Or _
      m_poi(wenti_data.point_no(3)).data(0).is_change Then
   Call incenter1(wenti_data.point_no(1), _
  wenti_data.point_no(2), _
     wenti_data.point_no(3), _
       wenti_data.point_no(0), _
         m_Circ(wenti_data.point_no(10)).data(0).data0.radii, _
          wenti_data.point_no(4), wenti_data.point_no(5), _
           wenti_data.point_no(6), True)
         'poi(wenti_data.point_no(0)).data(0).is_change
         'poi(wenti_data.point_no(4)).data(0).is_change
         'poi(wenti_data.point_no(5)).data(0).is_change
         'poi(wenti_data.point_no(6)).data(0).is_change
        ' Call change_circle_(wenti_data.point_no(0), pointapi0)
  End If

End Function

Private Function change_picture23(wenti_data As wentitype) As Boolean
'□、□、□、□四点共圆
change_picture23 = True
Dim c_data0 As circle_data0_type
If m_poi(wenti_data.point_no(0)).data(0).is_change Or _
    m_poi(wenti_data.point_no(1)).data(0).is_change Or _
     m_poi(wenti_data.point_no(2)).data(0).is_change Then
      Call read_three_circle0(wenti_data.point_no(12), _
         m_poi(wenti_data.point_no(0)).data(0).data0.coordinate, _
          m_poi(wenti_data.point_no(1)).data(0).data0.coordinate, _
           m_poi(wenti_data.point_no(2)).data(0).data0.coordinate, _
            m_Circ(wenti_data.point_no(12)).data(0).data0.c_coord, _
             m_Circ(wenti_data.point_no(12)).data(0).data0.radii, 0)
        
End If

End Function

Private Function change_picture29(wenti_data As wentitype) As Boolean
change_picture29 = True
t_coord.X = _
 (m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.X + _
  m_poi(wenti_data.point_no(2)).data(0).data0.coordinate.X) / 2
t_coord.Y = _
 (m_poi(wenti_data.point_no(1)).data(0).data0.coordinate.Y + _
  m_poi(wenti_data.point_no(2)).data(0).data0.coordinate.Y) / 2
 Call set_point_coordinate(wenti_data.point_no(3), t_coord, True)
End Function

Public Sub draw_item0(it0 As item0_data_type, condition_or_conclusion As Byte)
If it0.sig = "*" Or it0.sig = "/" Or it0.sig = "+" Or it0.sig = "-" Then
 Call draw_element_of_item0(it0.poi(0), it0.poi(1), condition_or_conclusion)
 Call draw_element_of_item0(it0.poi(2), it0.poi(3), condition_or_conclusion)
ElseIf it0.sig = "~" Then
 Call draw_element_of_item0(it0.poi(0), it0.poi(1), condition_or_conclusion)
End If
End Sub
Public Sub draw_angle(p1%, p2%, p3%, condition_or_conclusion As Byte)
Dim color As Byte
If condition_or_conclusion = condition Then
 color = condition_color
Else
 color = conclusion_color
End If
 If condition_or_conclusion = conclusion Then
 Call line_number(p1%, p2%, pointapi0, pointapi0, _
                  depend_condition(0, 0), depend_condition(0, 0), _
                  condition_or_conclusion, color, 1, 0)
 Call line_number(p2%, p3%, pointapi0, pointapi0, _
                  depend_condition(0, 0), depend_condition(0, 0), _
                  condition_or_conclusion, color, 1, 0)
 Else
 Call line_number(p1%, p2%, pointapi0, pointapi0, _
                  depend_condition(point_, p1%), depend_condition(point_, p2%), _
                  condition_or_conclusion, color, 1, 0)
 Call line_number(p2%, p3%, pointapi0, pointapi0, _
                  depend_condition(point_, p2%), depend_condition(point_, p3%), _
                  condition_or_conclusion, color, 1, 0)
 End If
End Sub
Public Sub draw_element_of_item0(ByVal p1%, ByVal p2%, condition_or_conclusion As Byte)
Dim color As Byte
If condition_or_conclusion = condition Then
   color = condition_color
Else
   color = conclusion_color
End If
 If p1% > 0 And p2% > 0 Then
    If condition_or_conclusion = conclusion Then
    Call line_number(p1%, p2%, pointapi0, pointapi0, _
         depend_condition(0, 0), depend_condition(0, 0), _
         condition_or_conclusion, color, 1, 0)
    Else
    Call line_number(p1%, p2%, pointapi0, pointapi0, _
                     depend_condition(point_, p1%), depend_condition(point_, p2%), _
                     condition_or_conclusion, color, 1, 0)
    End If
 ElseIf p2% > -6 Then
    Call draw_angle(angle(p1%).data(0).poi(0), angle(p1%).data(0).poi(1), _
                    angle(p1%).data(0).poi(2), condition_or_conclusion)
 End If
End Sub
Public Sub put_point_to_circle(p_coord As POINTAPI, c As circle_data0_type, out_coord As POINTAPI, out_p%)
Dim r!
         t_coord1 = minus_POINTAPI(p_coord, c.c_coord)
         r! = abs_POINTAPI(t_coord1)
  out_coord.X = c.c_coord.X + t_coord1.X * c.radii / r!
  out_coord.Y = c.c_coord.Y + t_coord1.Y * c.radii / r!
   If out_p% > 0 Then
     Call set_point_coordinate(out_p%, out_coord, True)
   End If
End Sub
Public Function change_picture_4(wenti_data As wentitype, _
                                     change_element As condition_type, t_cp%) As Boolean
change_picture_4 = True
If wenti_data.no_ = -50 Or wenti_data.no_ = -501 Or _
     wenti_data.no_ = -502 Then
    wenti_data.no_ = -50
     change_picture_4 = change_picture_50(wenti_data, change_element, t_cp%)
ElseIf wenti_data.no_ = -51 Then
     change_picture_4 = change_picture_51(wenti_data)
ElseIf wenti_data.no_ = -52 Then
     change_picture_4 = change_picture_52_56(wenti_data)
End If
End Function
Public Function change_picture8_71(wenti_data As wentitype, num As Integer, change_element As condition_type, cp%)
Dim tc%
If change_pointr% Or (wenti_data.point_no(29) <= change_pointr% And _
             wenti_data.point_no(28) >= change_pointr%) Then
   tc% = wenti_data.circ(1)
If wenti_data.no = 8 Then
   m_Circ(tc%).data(0).data0.c_coord = m_poi(wenti_data.poi(1)).data(0).data0.coordinate '设置两点圆变化后的圆心坐标
   m_Circ(tc%).data(0).data0.radii = distance_of_two_POINTAPI( _
                            m_poi(wenti_data.poi(1)).data(0).data0.coordinate, _
                              m_poi(wenti_data.poi(2)).data(0).data0.coordinate) '计算圆的半径
   m_Circ(tc%).data(0).is_change = True '
   Call change_m_circle(tc%, depend_condition(0, 0)) '
ElseIf wenti_data = -71 Then
  m_Circ(tc%).data(0).data0.radii = circle_radii0(m_poi(wenti_data.poi(1)).data(0).data0.coordinate, _
                                       m_poi(wenti_data.poi(2)).data(0).data0.coordinate, _
                                        m_poi(wenti_data.poi(3)).data(0).data0.coordinate, _
                                         m_Circ(tc%).data(0).data0.c_coord) '计算三点圆的半径和圆心坐标
    m_Circ(tc%).data(0).is_change = True
   Call change_m_circle(tc%, depend_condition(0, 0))
End If
End If
End Function

Public Function change_picture8_7101(wenti_data As wentitype, num As Integer, change_pointr%, cp%)
Dim p_coord(2) As POINTAPI
Dim ty(1) As Integer
If change_pointr% Or (wenti_data.point_no(29) <= change_pointr% And _
             wenti_data.point_no(28) >= change_pointr%) Then
   p_coord(0) = mid_POINTAPI(m_poi(wenti_data.poi(2)).data(0).data0.coordinate, _
                     m_poi(wenti_data.poi(3)).data(0).data0.coordinate)
   If wenti_data.circ(2) > 0 Then
   Call inter_point_line_circle3(p_coord(0), False, _
                  m_poi(wenti_data.poi(2)).data(0).data0.coordinate, _
                   m_poi(wenti_data.poi(3)).data(0).data0.coordinate, _
                  m_Circ(wenti_data.circ(2)).data(0).data0, p_coord(1), 0, p_coord(2), 0, _
                   wenti_data.inter_set_point_type, True)
        If wenti_data.inter_set_point_type = 0 Then
           If distance_of_two_POINTAPI(p_coord(1), m_poi(wenti_data.poi(1)).data(0).data0.coordinate) < _
               distance_of_two_POINTAPI(p_coord(1), m_poi(wenti_data.poi(1)).data(0).data0.coordinate) Then
                ty(0) = 1
           Else
                ty(0) = 2
           End If
           ty(1) = 0
        Else
        ty(0) = wenti_data.inter_set_point_type Mod 3
        ty(1) = wenti_data.inter_set_point_type \ 3
        End If
        If ty(0) = 1 Then
             Call set_point_coordinate(wenti_data.poi(1), p_coord(1), True)
        ElseIf ty(0) = 2 Then
             Call set_point_coordinate(wenti_data.poi(1), p_coord(2), True)
        End If
   ElseIf wenti_data.line_no(1) > 0 Then
           Call inter_point_line_line2(p_coord(0), verti_, _
                  m_poi(wenti_data.poi(2)).data(0).data0.coordinate, _
                   m_poi(wenti_data.poi(3)).data(0).data0.coordinate, _
                  m_poi(m_lin(wenti_data.line_no(2)).data(0).data0.poi(0)).data(0).data0.coordinate, _
                   paral_, m_poi(m_lin(wenti_data.line_no(2)).data(0).data0.poi(0)).data(0).data0.coordinate, _
                     m_poi(m_lin(wenti_data.line_no(2)).data(0).data0.poi(1)).data(0).data0.coordinate, _
                    p_coord(1), wenti_data.poi(1), True, False)
        If wenti_data.inter_set_point_type = 0 Then
           ty(0) = 1
           ty(1) = 0
        Else
           ty(0) = wenti_data.inter_set_point_type Mod 3
           ty(1) = wenti_data.inter_set_point_type \ 3
        End If
   End If
        Call change_picture_0(wenti_data.poi(1), wenti_data.point_no(29), num)
   If m_poi(wenti_data.poi(4)).data(0).degree = 2 Then
       Call put_point_to_circle(m_poi(wenti_data.poi(4)).data(0).data0.coordinate, _
            m_Circ(wenti_data.circ(1)).data(0).data0, _
                       p_coord(1), wenti_data.poi(4))
       If wenti_data.inter_set_point_type = 0 Then
          Call C_display_wenti.set_m_inner_point_type(num, ty(0))
       End If
   ElseIf m_poi(wenti_data.poi(4)).data(0).degree = 1 Then
       If m_poi(wenti_data.poi(4)).data(0).parent.element(0).ty = line_ Then
          Call inter_point_line_circle3( _
               m_poi(m_lin(wenti_data.line_no(3)).data(0).data0.poi(0)).data(0).data0.coordinate, _
                True, m_poi(m_lin(wenti_data.line_no(3)).data(0).data0.poi(0)).data(0).data0.coordinate, _
                 m_poi(m_lin(wenti_data.line_no(3)).data(0).data0.poi(1)).data(0).data0.coordinate, _
                  m_Circ(wenti_data.circ(1)).data(0).data0, p_coord(1), 0, p_coord(2), 0, 0, True)
       ElseIf m_poi(wenti_data.poi(4)).data(0).parent.element(0).ty = circle_ Then
          Call inter_point_circle_circle_( _
                m_Circ(wenti_data.circ(3)).data(0).data0, _
                  m_Circ(wenti_data.circ(1)).data(0).data0, p_coord(1), 0, p_coord(2), 0, 0, 0, True)
       End If
               If ty(1) = 0 Then
                If distance_of_two_POINTAPI(p_coord(1), m_poi(wenti_data.poi(4)).data(0).data0.coordinate) < _
                   distance_of_two_POINTAPI(p_coord(2), m_poi(wenti_data.poi(4)).data(0).data0.coordinate) Then
                  ty(1) = 1
                Else
                  ty(1) = 2
                End If
                  Call C_display_wenti.set_m_inner_point_type(num, ty(0) + 3 * ty(1))
               End If
               If ty(1) = 1 Then
                  Call set_point_coordinate(wenti_data.poi(4), p_coord(1), True)
               Else
                  Call set_point_coordinate(wenti_data.poi(4), p_coord(2), True)
               End If
   End If
           cp% = wenti_data.poi(4)
End If
End Function
Public Sub change_m_point(ByVal new_point_no%, Optional is_first_time As Boolean = False)
'new_point_no%移动点,from_condition 移动点的来历,is_first_time 第一次移点
Dim i%, j%, t_line%
Dim tp(1) As Integer
Dim r!
Dim t_coord(1) As POINTAPI
Dim t_in_point(10) As Integer
If m_poi(new_point_no%).data(0).is_change Then
    Exit Sub
Else
   m_poi(new_point_no%).data(0).is_change = True
End If
      'If m_Poi(new_point_no%).is_first_time And m_poi(new_point_no%).data(0).parent.co_degree = 1 Then '首次移动，引用的是鼠标点的位置（引用的点的序号，是否要做这样的预处理？）
       If m_poi(new_point_no%).data(0).parent.co_degree = 1 Then '自由度=1，非自由点（后续作图产生的点），
          If m_poi(new_point_no%).data(0).parent.element(1).ty = line_ Then '如果点在直线上,因为种种原因 degree=1 所以last_parent=1
           t_line% = m_poi(new_point_no%).data(0).parent.element(1).no
           If is_first_time = True Then
            If m_lin(t_line%).data(0).data0.depend_poi(0) <> new_point_no% And m_lin(t_line%).data(0).data0.depend_poi(1) <> new_point_no% Then
             Call orthofoot1(m_poi(new_point_no%).data(0).data0.coordinate, _
               second_end_point_coordinate(t_line%), _
                  m_poi(m_lin(t_line%).data(0).data0.depend_poi(0)).data(0).data0.coordinate, _
                    m_poi(new_point_no%).data(0).data0.coordinate, new_point_no%)
                     m_poi(new_point_no%).data(0).is_change = True '将移动点，投影到直线上，设置移动点的移动属性真
                '点与直线端点的相对位置（比值）
                m_poi(new_point_no%).data(0).parent.ratio = get_ratio_of_point_on_line(new_point_no%, _
                     m_poi(new_point_no%).data(0).parent.element(1).no, _
                       m_poi(new_point_no%).data(0).parent.related_point(0), _
                         m_poi(new_point_no%).data(0).parent.related_point(1), _
                          m_poi(new_point_no%).data(0).parent.related_point(2))
            '*******************************************************************************************************
             ' Call arrange_move_point_on_line(new_point_no%, t_line%)
              End If
            Else 'is_first_tiem=false
               m_poi(new_point_no%).data(0).data0.coordinate = _
                 get_coordinate_of_point_on_line(new_point_no%, m_poi(new_point_no%).data(0).parent.element(1).no)
            End If
      '*************************************************************************************************************
         ElseIf m_poi(new_point_no%).data(0).parent.element(1).ty = circle_ And _
                 new_point_no% <> m_Circ(m_poi(new_point_no%).data(0).parent.element(1).no).data(0).data0.center Then  '点在圆上
           m_poi(new_point_no%).data(0).data0.coordinate = _
             add_POINTAPI(m_Circ(m_poi(new_point_no%).data(0).parent.element(1).no).data(0).data0.c_coord, _
                  time_POINTAPI_by_number(minus_POINTAPI(m_poi(new_point_no%).data(0).data0.coordinate, _
                    m_Circ(m_poi(new_point_no%).data(0).parent.element(1).no).data(0).data0.c_coord), _
                     m_Circ(m_poi(new_point_no%).data(0).parent.element(1).no).data(0).data0.radii / _
                      (distance_of_two_POINTAPI(m_poi(new_point_no%).data(0).data0.coordinate, _
                    m_Circ(m_poi(new_point_no%).data(0).parent.element(1).no).data(0).data0.c_coord))))
            '将移动点从圆心投影到圆上
         End If
        ElseIf m_poi(new_point_no%).data(0).parent.co_degree = 2 Then
         If m_poi(new_point_no%).data(0).parent.element(1).ty = line_ And m_poi(new_point_no%).data(0).parent.element(2).ty = line_ Then
          '两直线交点
           Call calculate_line_line_intersect_point( _
           m_poi(m_lin(m_poi(new_point_no%).data(0).parent.element(1).no).data(0).data0.depend_poi(0)).data(0).data0.coordinate, _
           second_end_point_coordinate(m_poi(new_point_no%).data(0).parent.element(1).no), _
           m_poi(m_lin(m_poi(new_point_no%).data(0).parent.element(2).no).data(0).data0.depend_poi(0)).data(0).data0.coordinate, _
           second_end_point_coordinate(m_poi(new_point_no%).data(0).parent.element(2).no), _
           m_poi(new_point_no%).data(0).data0.coordinate, False)
         ElseIf m_poi(new_point_no%).data(0).parent.element(1).ty = line_ And m_poi(new_point_no%).data(0).parent.element(2).ty = circle_ Then
           Call inter_point_line_circle(m_poi(new_point_no%).data(0).parent.element(1).no, _
                  m_poi(new_point_no%).data(0).parent.element(2).no, m_poi(new_point_no%).data(0).data0.coordinate, _
                   new_point_no%, False, False, m_poi(new_point_no%).data(0).parent.inter_type)
         ElseIf m_poi(new_point_no%).data(0).parent.element(1).ty = circle_ And m_poi(new_point_no%).data(0).parent.element(2).ty = circle_ Then
            If m_poi(new_point_no%).data(0).parent.inter_type = tangent_point_ And m_poi(new_point_no%).data(0).parent.related_point(0) > 0 Then
              Call change_tangent_line_for_two_circle(m_poi(new_point_no%).data(0).parent.related_point(0))
            Else
            Call inter_point_circle_circle_for_change(m_poi(new_point_no%).data(0).parent.element(1).no, _
                     m_poi(new_point_no%).data(0).parent.element(2).no, m_poi(new_point_no%).data(0).data0.coordinate, _
                         m_poi(new_point_no%).data(0).parent.inter_type)
            End If
         ElseIf m_poi(new_point_no%).data(0).parent.element(1).ty = point_ And m_poi(new_point_no%).data(0).parent.element(2).ty = circle_ Then
          If m_poi(new_point_no%).data(0).parent.inter_type = new_point_on_circle_circle12 Then
          m_poi(new_point_no%).data(0).data0.coordinate = _
           inter_point_circle_circle_by_pointapi(m_Circ(m_poi(new_point_no%).data(0).parent.element(2).no).data(0).data0.c_coord, _
             m_Circ(m_poi(new_point_no%).data(0).parent.element(2).no).data(0).data0.radii, _
               mid_POINTAPI(m_poi(m_poi(new_point_no%).data(0).parent.element(1).no).data(0).data0.coordinate, _
                m_Circ(m_poi(new_point_no%).data(0).parent.element(2).no).data(0).data0.c_coord), _
                distance_of_two_POINTAPI(m_poi(m_poi(new_point_no%).data(0).parent.element(1).no).data(0).data0.coordinate, _
                  m_Circ(m_poi(new_point_no%).data(0).parent.element(2).no).data(0).data0.c_coord) / 2, _
                   t_coord(0), t_coord(1), new_point_on_circle_circle12)
          ElseIf m_poi(new_point_no%).data(0).parent.inter_type = new_point_on_circle_circle21 Then
              m_poi(new_point_no%).data(0).data0.coordinate = _
           inter_point_circle_circle_by_pointapi(m_Circ(m_poi(new_point_no%).data(0).parent.element(2).no).data(0).data0.c_coord, _
             m_Circ(m_poi(new_point_no%).data(0).parent.element(2).no).data(0).data0.radii, _
               mid_POINTAPI(m_poi(m_poi(new_point_no%).data(0).parent.element(1).no).data(0).data0.coordinate, _
                m_Circ(m_poi(new_point_no%).data(0).parent.element(2).no).data(0).data0.c_coord), _
                distance_of_two_POINTAPI(m_poi(m_poi(new_point_no%).data(0).parent.element(1).no).data(0).data0.coordinate, _
                  m_Circ(m_poi(new_point_no%).data(0).parent.element(2).no).data(0).data0.c_coord) / 2, _
                   t_coord(0), t_coord(1), new_point_on_circle_circle21)
          End If
         ElseIf m_poi(new_point_no%).data(0).parent.element(1).ty = line_ Then    '
                m_poi(new_point_no%).data(0).data0.coordinate = _
                 get_coordinate_of_point_on_line(new_point_no%, m_poi(new_point_no%).data(0).parent.element(1).no)
        ElseIf m_poi(new_point_no%).data(0).parent.element(2).ty = line_ Then
                m_poi(new_point_no%).data(0).data0.coordinate = _
                 get_coordinate_of_point_on_line(new_point_no%, m_poi(new_point_no%).data(0).parent.element(2).no)
         End If
        End If
      '‘ElseIf is_first_time = False Then
      '不是第一次移动点
      '****************************************************************************
      '移动不是直线的生成点，移动生成点—》变化直线—》变化直线上的非生成点
       For i% = 1 To m_poi(new_point_no%).data(0).in_line(0)
          t_line% = m_poi(new_point_no%).data(0).in_line(i%)
               Call arrange_move_point_on_line(new_point_no%, t_line%)
       Next i%
       '*********************************************************************************************************
            '自由度=0，非自由点（后续作图产生的点）
            For i% = 1 To m_poi(new_point_no%).data(0).sons.last_son '对移动点的每一个后继几何元素作处理
             If m_poi(new_point_no%).data(0).sons.son(i%).ty = point_ Then
                Call change_m_point(m_poi(new_point_no%).data(0).sons.son(i%).no)
             ElseIf m_poi(new_point_no%).data(0).sons.son(i%).ty = line_ Then
               '后继是直线
               tp(0) = m_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).data0.poi(0)
                tp(1) = m_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).data0.poi(1)
                 If new_point_no% = tp(0) Or new_point_no% = tp(1) Then
                     If compare_two_point(m_poi(tp(0)).data(0).data0.coordinate, m_poi(tp(1)).data(0).data0.coordinate, _
                          0, 0, 2) = -1 Then '改变了生成点的顺序
                        Call exchange_two_integer(m_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).data0.poi(0), _
                              m_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).data0.poi(1)) '交换生成点的顺序
                     '****************************************************************************************************
                          For j% = 1 To m_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).data0.in_point(0)
                              '复制
                                t_in_point(j%) = m_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).data0.in_point(j%)
                          Next j%
                          For j% = 1 To m_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).data0.in_point(0)
                             '反序复制回原数据库
                               m_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).data0.in_point(j%) = _
                                   t_in_point(m_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).data0.in_point(0) + 1 - j%)
                          Next j%
                      '*******************************************************************************************************
                          '重设直线端点坐标
                          m_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).data0.end_point_coord(0) = _
                                 m_poi(m_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).data0.in_point(1)).data(0).data0.coordinate
                          m_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).data0.end_point_coord(1) = _
                                 m_poi(m_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).data0.in_point( _
                                   m_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).data0.in_point(0))).data(0).data0.coordinate
                       '***********************************************************************************************************
                     '改变全部点的顺序
                      End If
                 End If
                  'm_lin(m_poi(new_point_no%).data(0).sons.son(i%).no).data(0).is_change = True
                 Call change_m_line(m_poi(new_point_no%).data(0).sons.son(i%).no)
                '变化直线
             ElseIf m_poi(new_point_no%).data(0).sons.son(i%).ty = circle_ Then
                '后继是圆
                Call change_m_circle(m_poi(new_point_no%).data(0).sons.son(i%).no, depend_condition(point_, new_point_no%))
                '变化圆
             ElseIf m_poi(new_point_no%).data(0).sons.son(i%).ty = wenti_cond_ Then
                '后继是作图语句
                 Call change_picture(m_poi(new_point_no%).data(0).sons.son(i%).no, depend_condition(point_, new_point_no%))
                 '变化语句
             ElseIf m_poi(new_point_no%).data(0).sons.son(i%).ty = epolygon_ Then
                 Call change_epolygon(m_poi(new_point_no%).data(0).sons.son(i%).no)
             End If
      Next i%
 '*************************************************************************************************
       Call add_next_change_element(point_, new_point_no%)
       If is_first_time Then
        Call draw_change_picture
       End If
       'Call C_display_picture.draw_point(new_point_no%) '重画移动点
        'm_poi(new_point_no%).data(0).is_change = False '撤销移动属性
End Sub

Public Sub change_m_line(ByVal line_no%)
Dim i%
   'If m_lin(line_no%).data(0).is_change = False Then
   'If m_poi(m_lin(line_no%).data(0).data0.poi(0)).data(0).is_change Then
   ' m_lin(line_no%).data(0).data0.end_point_coord(0) = _
      m_poi(m_lin(line_no%).data(0).data0.poi(0)).data(0).data0.coordinate
   'End If
   'If m_poi(m_lin(line_no%).data(0).data0.poi(1)).data(0).is_change Then '如果是射线，不能由端点序号确定端点坐标
   ' m_lin(line_no%).data(0).data0.end_point_coord(1) = _
      m_poi(m_lin(line_no%).data(0).data0.poi(1)).data(0).data0.coordinate
   'End If
   'm_lin(line_no%).data(0).is_change = True
   'End If
   If m_lin(line_no%).data(0).is_change = 255 Then '阻止循环调用
      Exit Sub
   Else
      m_lin(line_no%).data(0).is_change = 255
   End If
   If m_lin(line_no%).data(0).parent.inter_type = paral_ Then
      m_lin(line_no%).data(0).data0.depend_poi1_coord = _
         add_POINTAPI(m_poi(m_lin(line_no%).data(0).data0.depend_poi(0)).data(0).data0.coordinate, _
           minus_POINTAPI(second_end_point_coordinate(m_lin(line_no%).data(0).parent.element(2).no), _
             m_poi(m_lin(m_lin(line_no%).data(0).parent.element(2).no).data(0).data0.depend_poi(0)).data(0).data0.coordinate))
   ElseIf m_lin(line_no%).data(0).parent.inter_type = verti_ Then
      m_lin(line_no%).data(0).data0.depend_poi1_coord = _
         add_POINTAPI(m_poi(m_lin(line_no%).data(0).data0.depend_poi(0)).data(0).data0.coordinate, _
           verti_POINTAPI(minus_POINTAPI(second_end_point_coordinate(m_lin(line_no%).data(0).parent.element(2).no), _
             m_poi(m_lin(m_lin(line_no%).data(0).parent.element(2).no).data(0).data0.depend_poi(0)).data(0).data0.coordinate)))
  End If
   For i% = 1 To m_lin(line_no%).data(0).sons.last_son
      If m_lin(line_no%).data(0).sons.son(i%).ty = point_ Then '直线上添加的点，按比例计算坐标
             'm_poi(m_lin(line_no%).data(0).sons.son(i%).no).data(0).data0.coordinate = _
                 add_POINTAPI(m_poi(m_lin(line_no%).data(0).data0.poi(1)).data(0).data0.coordinate, _
                   time_POINTAPI_by_number(minus_POINTAPI(m_poi(m_lin(line_no%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                     m_poi(m_lin(line_no%).data(0).data0.poi(1)).data(0).data0.coordinate), _
                        m_poi(m_lin(line_no%).data(0).sons.son(i%).no).data(0).parent.ratio))
            'm_poi(m_lin(line_no%).data(0).sons.son(i%).no).data(0).is_change = True
            Call change_m_point(m_lin(line_no%).data(0).sons.son(i%).no)
      ElseIf m_lin(line_no%).data(0).sons.son(i%).ty = line_ Then
       Call change_m_line(m_lin(line_no%).data(0).sons.son(i%).no)
      ElseIf m_lin(line_no%).data(0).sons.son(i%).ty = circle_ Then
       Call change_m_circle(m_lin(line_no%).data(0).sons.son(i%).no, depend_condition(line_, line_no%))
      ElseIf m_lin(line_no%).data(0).sons.son(i%).ty = wenti_cond_ Then
       Call change_picture(m_lin(line_no%).data(0).sons.son(i%).no, depend_condition(line_, line_no%))
      End If
    Next i%
    Call add_next_change_element(line_, line_no%)
      ' Call C_display_picture.re_draw_line(line_no%)
End Sub

Public Sub change_m_circle(ByVal circle_no%, from_condition As condition_type)
Dim i%
If m_Circ(circle_no%).data(0).is_change = False Then
If from_condition.ty = Ratio_for_measure_ Then '标尺发生变化
  m_Circ(circle_no%).data(0).data0.radii = m_Circ(circle_no%).data(0).data0.real_radii * _
      Ratio_for_measure.Ratio_for_measure '计算圆的显示半径
Else '有关圆的点发生变化，计算半径
If m_Circ(circle_no%).data(0).data0.center = 0 And _
    (m_poi(m_Circ(circle_no%).data(0).data0.in_point(1)).data(0).is_change Or _
       m_poi(m_Circ(circle_no%).data(0).data0.in_point(2)).data(0).is_change Or _
        m_poi(m_Circ(circle_no%).data(0).data0.in_point(3)).data(0).is_change) And _
         m_Circ(circle_no%).data(0).data0.real_radii = 0 Then
    '三点圆
    m_Circ(circle_no%).data(0).data0.radii = _
        circle_radii0(m_poi(m_Circ(circle_no%).data(0).data0.in_point(1)).data(0).data0.coordinate, _
                       m_poi(m_Circ(circle_no%).data(0).data0.in_point(2)).data(0).data0.coordinate, _
                        m_poi(m_Circ(circle_no%).data(0).data0.in_point(3)).data(0).data0.coordinate, _
                         m_Circ(circle_no%).data(0).data0.c_coord)
         m_Circ(circle_no%).data(0).is_change = True
ElseIf m_Circ(circle_no%).data(0).data0.center > 0 Then
        If (m_poi(m_Circ(circle_no%).data(0).data0.center).data(0).is_change Or _
             m_poi(m_Circ(circle_no%).data(0).data0.in_point(1)).data(0).is_change) Then
               m_Circ(circle_no%).data(0).data0.c_coord = _
                m_poi(m_Circ(circle_no%).data(0).data0.center).data(0).data0.coordinate
              If (m_Circ(circle_no%).data(0).data0.real_radii = 0 And _
                 m_Circ(circle_no%).data(0).parent.inter_type <> length_depended_by_two_points_) Then
                 m_Circ(circle_no%).data(0).data0.radii = distance_of_two_POINTAPI( _
                   m_poi(m_Circ(circle_no%).data(0).data0.center).data(0).data0.coordinate, _
                    m_poi(m_Circ(circle_no%).data(0).data0.in_point(1)).data(0).data0.coordinate)
              End If
         ElseIf m_Circ(circle_no%).data(0).parent.inter_type = length_depended_by_two_points_ And _
               (m_poi(m_Circ(circle_no%).data(0).parent.related_point(0)).data(0).is_change = True Or _
                 m_poi(m_Circ(circle_no%).data(0).parent.related_point(1)).data(0).is_change = True) Then  '圆的半径有两个指定点确定
           m_Circ(circle_no%).data(0).data0.radii = distance_of_two_POINTAPI( _
                   m_poi(m_Circ(circle_no%).data(0).parent.related_point(0)).data(0).data0.coordinate, _
                    m_poi(m_Circ(circle_no%).data(0).parent.related_point(1)).data(0).data0.coordinate)
           m_Circ(circle_no%).data(0).data0.real_radii = m_Circ(circle_no%).data(0).data0.radii
        End If
        '       If m_Circ(circle_no%).data(0).parent.element(1).ty = point_ And _
             m_Circ(circle_no%).data(0).parent.element(1).no = m_Circ(circle_no%).data(0).data0.in_point(1) Then
        'm_Circ(circle_no%).data(0).data0.radii = distance_of_two_POINTAPI( _
          m_poi(m_Circ(circle_no%).data(0).data0.center).data(0).data0.coordinate, _
            m_poi(m_Circ(circle_no%).data(0).data0.in_point(1)).data(0).data0.coordinate)
        End If
     If m_Circ(circle_no%).data(0).parent.inter_type = circle_radio_relate_by_two_point_ Then
       m_Circ(circle_no%).data(0).data0.radii = abs_POINTAPI(minus_POINTAPI( _
        m_poi(m_Circ(circle_no%).data(0).parent.related_point(0)).data(0).data0.coordinate, _
         m_poi(m_Circ(circle_no%).data(0).parent.related_point(1)).data(0).data0.coordinate))
     End If
End If
End If
   m_Circ(circle_no%).data(0).is_change = True
'*************************************************************************************************
For i% = 1 To m_Circ(circle_no%).data(0).sons.last_son
    If m_Circ(circle_no%).data(0).sons.son(i%).ty = point_ Then
          '(from_condition.ty <> point_ Or from_condition.no <> m_Circ(circle_no%).data(0).sons.son(i%).no) Then
          ' m_poi(m_Circ(circle_no%).data(0).sons.son(i%).no).data(0).data0.coordinate = _
             add_POINTAPI(m_Circ(circle_no%).data(0).data0.c_coord, _
                  time_POINTAPI_by_number(minus_POINTAPI(m_poi(m_Circ(circle_no%).data(0).sons.son(i%).no).data(0).data0.coordinate, _
                    m_Circ(circle_no%).data(0).data0.c_coord), _
                     m_Circ(circle_no%).data(0).data0.radii / _
                      (distance_of_two_POINTAPI(m_poi(m_Circ(circle_no%).data(0).sons.son(i%).no).data(0).data0.coordinate, _
                    m_Circ(circle_no%).data(0).data0.c_coord))))
           'm_poi(m_Circ(circle_no%).data(0).sons.son(i%).no).data(0).is_change = True
           Call change_m_point(m_Circ(circle_no%).data(0).sons.son(i%).no)
    ElseIf m_Circ(circle_no%).data(0).sons.son(i%).ty = line_ And _
       (from_condition.ty <> line_ Or from_condition.no <> m_Circ(circle_no%).data(0).sons.son(i%).no) Then
       Call change_m_line(m_Circ(circle_no%).data(0).sons.son(i%).no)
    ElseIf m_Circ(circle_no%).data(0).sons.son(i%).ty = circle_ And _
       (from_condition.ty <> circle_ Or from_condition.no <> m_Circ(circle_no%).data(0).sons.son(i%).no) Then
       Call change_m_circle(m_Circ(circle_no%).data(0).sons.son(i%).no, depend_condition(circle_, circle_no%))
    ElseIf m_Circ(circle_no%).data(0).sons.son(i%).ty = wenti_cond_ Then
       If from_condition.ty <> wenti_cond_ Or from_condition.no <> m_Circ(circle_no%).data(0).sons.son(i%).no Then
        Call change_picture(m_Circ(circle_no%).data(0).sons.son(i%).no, depend_condition(circle_, circle_no), 0)
       End If
    End If
Next i%
 Call add_next_change_element(circle_, circle_no%)
 'Call C_display_picture.draw_circle(circle_no%, 0, 0)
 
  '   m_Circ(circle_no%).data(0).is_change = False
End Sub

Public Sub change_tangent_line_for_two_circle(ByVal tangent_line_no%)
Dim r&, D1&, D2&
Dim i%, l%, out_c1%, out_c2%
Dim sr&, temp_k1!, temp_k2!
Dim co!, si!
Dim ty As Byte
Dim p_coord(1) As POINTAPI
Dim p As POINTAPI
Call set_tangent_line_for_two_circle(tangent_line(tangent_line_no%).data(0).circ(0), _
                         tangent_line(tangent_line_no%).data(0).circ(1), 0, _
                           tangent_line_no%, tangent_line(tangent_line_no%).tangent_type)
     Exit Sub
     
If tangent_line(tangent_line_no%).tangent_type < 5 Then
 ty = tangent_line(tangent_line_no%).tangent_type
If tangent_line(tangent_line_no%).data(0).ele(0).ty = circle_ Then
out_c1% = tangent_line(tangent_line_no%).data(0).ele(0).no
End If
If tangent_line(tangent_line_no%).data(0).ele(1).ty = circle_ Then
out_c2% = tangent_line(tangent_line_no%).data(0).ele(1).no
End If
If out_c1% > out_c2% Then
  Call exchange_two_integer(out_c1%, out_c2%)
End If
If tangent_line(tangent_line_no%).tangent_type < 5 Then

r& = (m_Circ(out_c1%).data(0).data0.c_coord.X - _
       m_Circ(out_c2%).data(0).data0.c_coord.X) ^ 2 + _
     (m_Circ(out_c1%).data(0).data0.c_coord.Y - _
       m_Circ(out_c2%).data(0).data0.c_coord.Y) ^ 2
        sr& = sqr(r&)
         temp_k1! = CSng((m_Circ(out_c1%).data(0).data0.c_coord.X - _
                     m_Circ(out_c2%).data(0).data0.c_coord.X) / sr&)
         temp_k2! = CSng((m_Circ(out_c1%).data(0).data0.c_coord.Y - _
                     m_Circ(out_c2%).data(0).data0.c_coord.Y) / sr&)
D1& = m_Circ(out_c2%).data(0).data0.radii - m_Circ(out_c1%).data(0).data0.radii
D2& = m_Circ(out_c2%).data(0).data0.radii + m_Circ(out_c1%).data(0).data0.radii
'****************************************************************************************

If sr& > D2& Then '内公切线
         co! = CSng(D2& / sr&)
         si! = sqr(r& - D2& ^ 2) / sr&
If ty = 1 Then
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.X = _
        m_Circ(out_c2%).data(0).data0.c_coord.X + _
        m_Circ(out_c2%).data(0).data0.radii * (co! * temp_k1! - si! * temp_k2!)
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.Y = _
        m_Circ(out_c2%).data(0).data0.c_coord.Y + _
        m_Circ(out_c2%).data(0).data0.radii * (si! * temp_k1! + co! * temp_k2!)
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.X = _
        m_Circ(out_c1%).data(0).data0.c_coord.X - _
        m_Circ(out_c1%).data(0).data0.radii * (co! * temp_k1! - si! * temp_k2!)
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.Y = _
        m_Circ(out_c1%).data(0).data0.c_coord.Y - _
        m_Circ(out_c1%).data(0).data0.radii * (si! * temp_k1! + co! * temp_k2!)
'          Call set_tangent_line_data(p_coord(0), p_coord(1), _
                depend_condition(circle_, out_c1%), depend_condition(circle_, out_c2%), 1, 1)
End If
If ty = 2 Then
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.X = _
        m_Circ(out_c2%).data(0).data0.c_coord.X + _
        m_Circ(out_c2%).data(0).data0.radii * (co! * temp_k1! + si! * temp_k2!)
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.Y = _
        m_Circ(out_c2%).data(0).data0.c_coord.Y + _
        m_Circ(out_c2%).data(0).data0.radii * (-si! * temp_k1! + co! * temp_k2!)

m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.X = _
        m_Circ(out_c1%).data(0).data0.c_coord.X - _
        m_Circ(out_c1%).data(0).data0.radii * (co! * temp_k1! + si! * temp_k2!)
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.Y = _
        m_Circ(out_c1%).data(0).data0.c_coord.Y - _
        m_Circ(out_c1%).data(0).data0.radii * (-si! * temp_k1! + co! * temp_k2!)
'  Call set_tangent_line_data(p_coord(0), p_coord(1), _
                depend_condition(circle_, out_c1%), depend_condition(circle_, out_c2%), 1, 2)
End If
End If

If -3 < D1& And D1& < 3 Then '等圆，外公切线
If ty = 3 Then
p.X = temp_k2! * m_Circ(out_c1%).data(0).data0.radii
p.Y = -temp_k1! * m_Circ(out_c1%).data(0).data0.radii
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.X = _
            m_Circ(out_c1%).data(0).data0.c_coord.X - p.X
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.Y = _
            m_Circ(out_c1%).data(0).data0.c_coord.Y - p.Y
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.X = _
            m_Circ(out_c2%).data(0).data0.c_coord.X - p.X
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.Y = _
            m_Circ(out_c2%).data(0).data0.c_coord.Y - p.Y
ElseIf ty = 4 Then
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.X = _
            m_Circ(out_c1%).data(0).data0.c_coord.X + p.X
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.Y = _
            m_Circ(out_c1%).data(0).data0.c_coord.Y + p.Y
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.X = _
            m_Circ(out_c2%).data(0).data0.c_coord.X + p.X
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.Y = _
            m_Circ(out_c2%).data(0).data0.c_coord.Y + p.Y
End If
Else
            
If sr& < D1& Then
  Exit Sub
ElseIf sr& >= Abs(D1&) Then  '连心线长大于两圆半径差，有外切公切线
 If D1& < 0 Then
    D1& = -D1&
    Call exchange_two_integer(out_c1%, out_c2%)
    temp_k1! = -temp_k1!
    temp_k2! = -temp_k2!
    If ty = 3 Then
      ty = 4
    ElseIf ty = 4 Then
      ty = 3
    End If
 End If
         co! = CSng(D1& / sr&)
         si! = sqr(r& - D1& ^ 2) / sr&
If ty = 3 Then
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.X = _
        m_Circ(out_c2%).data(0).data0.c_coord.X + _
        m_Circ(out_c2%).data(0).data0.radii * (co! * temp_k1! - si! * temp_k2!)
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.Y = _
        m_Circ(out_c2%).data(0).data0.c_coord.Y + _
        m_Circ(out_c2%).data(0).data0.radii * (si! * temp_k1! + co! * temp_k2!)
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.X = _
        m_Circ(out_c1%).data(0).data0.c_coord.X + _
        m_Circ(out_c1%).data(0).data0.radii * (co! * temp_k1! - si! * temp_k2!)
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.Y = _
        m_Circ(out_c1%).data(0).data0.c_coord.Y + _
        m_Circ(out_c1%).data(0).data0.radii * (si! * temp_k1! + co! * temp_k2!)
ElseIf ty = 4 Then
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.X = _
        m_Circ(out_c2%).data(0).data0.c_coord.X + _
        m_Circ(out_c2%).data(0).data0.radii * (co! * temp_k1! + si! * temp_k2!)
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.Y = _
        m_Circ(out_c2%).data(0).data0.c_coord.Y + _
        m_Circ(out_c2%).data(0).data0.radii * (-si! * temp_k1! + co! * temp_k2!)
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.X = _
        m_Circ(out_c1%).data(0).data0.c_coord.X + _
        m_Circ(out_c1%).data(0).data0.radii * (co! * temp_k1! + si! * temp_k2!)
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.Y = _
        m_Circ(out_c1%).data(0).data0.c_coord.Y + _
        m_Circ(out_c1%).data(0).data0.radii * (-si! * temp_k1! + co! * temp_k2!)
End If
End If
End If
End If
        'm_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).is_change = True
        'm_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).is_change = True
        Call change_m_point(tangent_line(tangent_line_no%).data(0).poi(0))
        Call change_m_point(tangent_line(tangent_line_no%).data(0).poi(1))
ElseIf tangent_line(tangent_line_no%).tangent_type < 8 Then
'*****************************************************************************************************
If tangent_line(tangent_line_no%).tangent_type = 5 Then
 For i% = 1 To m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.in_point(0)
  If m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.in_point(i%) = _
                        tangent_line(tangent_line_no%).data(0).poi(1) Then
  m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate = _
                 m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate
  m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate = _
                 add_POINTAPI(m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate, _
                  verti_POINTAPI(minus_POINTAPI( _
                  m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.c_coord, _
                   m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate)))
                   
                 Exit Sub
  End If
 Next i%
Else
r& = (m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.c_coord.X - _
       m_poi(tangent_line(tangent_line_no%).data(0).ele(1).no).data(0).data0.coordinate.X) ^ 2 + _
     (m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.c_coord.Y - _
       m_poi(tangent_line(tangent_line_no%).data(0).ele(1).no).data(0).data0.coordinate.Y) ^ 2
        sr& = sqr(r&)
         temp_k1! = CSng((m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.c_coord.X - _
                     m_poi(tangent_line(tangent_line_no%).data(0).ele(1).no).data(0).data0.coordinate.X) / sr&)
         temp_k2! = CSng((m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.c_coord.Y - _
                     m_poi(tangent_line(tangent_line_no%).data(0).ele(1).no).data(0).data0.coordinate.Y) / sr&)
If sr& > m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.radii Then
            
         co! = CSng(m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.radii / sr&)
         si! = sqr(r& - m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.radii ^ 2) / sr&
 If tangent_line(tangent_line_no%).tangent_type = 6 Then
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.X = _
          m_poi(tangent_line(tangent_line_no%).data(0).ele(1).no).data(0).data0.coordinate.X
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.Y = _
          m_poi(tangent_line(tangent_line_no%).data(0).ele(1).no).data(0).data0.coordinate.Y
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.X = _
          m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.c_coord.X - _
        m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.radii * (co! * temp_k1! - si! * temp_k2!)
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.Y = _
        m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.c_coord.Y - _
        m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.radii * (si! * temp_k1! + co! * temp_k2!)
 '         Call set_tangent_line_data(p_coord(0), p_coord(1), _
 '               depend_condition(circle_, in_circle_no%), depend_condition(point_, in_point_no%), 6)
Else
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.X = _
        m_poi(tangent_line(tangent_line_no%).data(0).ele(1).no).data(0).data0.coordinate.X
m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate.Y = _
        m_poi(tangent_line(tangent_line_no%).data(0).ele(1).no).data(0).data0.coordinate.Y
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.X = _
        m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.c_coord.X - _
        m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.radii * (co! * temp_k1! + si! * temp_k2!)
m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate.Y = _
        m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.c_coord.Y - _
        m_Circ(tangent_line(tangent_line_no%).data(0).ele(0).no).data(0).data0.radii * (-si! * temp_k1! + co! * temp_k2!)
'  Call set_tangent_line_data(p_coord(0), p_coord(1), _
'                depend_condition(circle_, in_circle_no%), depend_condition(point_, in_point_no%), 7)
End If
End If
End If
'  m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).is_change = True
  m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).is_change = True
'       Call C_display_picture.draw_point(tangent_line(tangent_line_no%).data(0).poi(0)) '不能用change_m_point，否则有循环引用
  Call change_m_point(tangent_line(tangent_line_no%).data(0).poi(0))
End If
End Sub

Public Sub change_paral_or_verti_line(l1%, p1%, l2%, paral_or_verti_ As Integer)
If paral_or_verti_ = paral_ Then
   m_lin(l2%).data(0).data0.end_point_coord(0) = m_poi(p1%).data(0).data0.coordinate
   m_lin(l2%).data(0).data0.end_point_coord(1) = add_POINTAPI(m_poi(p1%).data(0).data0.coordinate, _
        minus_POINTAPI(m_lin(l1%).data(0).data0.end_point_coord(1), _
           m_lin(l1%).data(0).data0.end_point_coord(0)))
   'm_lin(l2%).data(0).is_change = True
   Call change_m_line(l2%)
ElseIf paral_or_verti_ = verti_ Then
   m_lin(l2%).data(0).data0.end_point_coord(0) = m_poi(p1%).data(0).data0.coordinate
   m_lin(l2%).data(0).data0.end_point_coord(1) = add_POINTAPI(m_poi(p1%).data(0).data0.coordinate, _
        verti_POINTAPI(minus_POINTAPI(m_lin(l1%).data(0).data0.end_point_coord(1), _
           m_lin(l1%).data(0).data0.end_point_coord(0))))
   'm_lin(l2%).data(0).is_change = True
   Call change_m_line(l2%)
End If
End Sub

Public Sub arrange_move_point_on_line(point_no%, t_line%)
Dim t_p%, jug%, i%
Dim need_redraw As Boolean
Dim p1%, pn%
Dim temp_color(1 To 9) As Byte
If point_no% > 0 Then
   p1% = m_lin(t_line%).data(0).data0.in_point(1)
   pn% = m_lin(t_line%).data(0).data0.in_point(m_lin(t_line%).data(0).data0.in_point(0))
         If (m_lin(t_line%).data(0).parent.element(1).ty = point_ And m_lin(t_line%).data(0).parent.element(1).no = point_no%) Or _
              (m_lin(t_line%).data(0).parent.element(1).ty = point_ And m_lin(t_line%).data(0).parent.element(2).no = point_no%) Then
             Exit Sub
         End If
         If m_lin(t_line%).data(0).data0.in_point(1) = point_no% Then '移动点是直线的第一个端点
              If compare_two_point(m_poi(m_lin(t_line%).data(0).data0.in_point(1)).data(0).data0.coordinate, _
                                 m_poi(m_lin(t_line%).data(0).data0.in_point(2)).data(0).data0.coordinate, _
                                  m_lin(t_line%).data(0).data0.poi(0), _
                                   m_lin(t_line%).data(0).data0.poi(1), 4) = -1 Then  '直线的第一点和第二点顺序改变
                    Call exchange_two_integer(m_lin(t_line%).data(0).data0.in_point(1), _
                                         m_lin(t_line%).data(0).data0.in_point(2)) '交换第一点和第二点
              End If
                              m_lin(t_line%).data(0).is_change = 255 '不论是否改变了直线上点的排列顺序，直线的端点坐标都发生变化
                              m_lin(t_line%).data(0).data0.end_point_coord(0) = _
                                m_poi(m_lin(t_line%).data(0).data0.in_point(1)).data(0).data0.coordinate
          ElseIf m_lin(t_line%).data(0).data0.in_point(m_lin(t_line%).data(0).data0.in_point(0)) = point_no% Then '移动点是最后一点
                If compare_two_point(m_poi(m_lin(t_line%).data(0).data0.in_point(m_lin(t_line%).data(0).data0.in_point(0) - 1)).data(0).data0.coordinate, _
                    m_poi(m_lin(t_line%).data(0).data0.in_point(m_lin(t_line%).data(0).data0.in_point(0))).data(0).data0.coordinate, _
                       m_lin(t_line%).data(0).data0.poi(0), _
                        m_lin(t_line%).data(0).data0.poi(1), 4) = -1 Then
                    Call exchange_two_integer(m_lin(t_line%).data(0).data0.in_point(m_lin(t_line%).data(0).data0.in_point(0) - 1), _
                                         m_lin(t_line%).data(0).data0.in_point(m_lin(t_line%).data(0).data0.in_point(0)))
                                
                End If
                               m_lin(t_line%).data(0).is_change = 255
                    m_lin(t_line%).data(0).data0.end_point_coord(1) = _
                                 m_poi(m_lin(t_line%).data(0).data0.in_point(m_lin(t_line%).data(0).data0.in_point(0))).data(0).data0.coordinate
                               'Call C_display_picture.re_draw_line(t_line%)
         ElseIf is_point_in_line3(point_no%, m_lin(t_line%).data(0).data0, t_p%) Then '读取移动点在直线上的序号（t_p%,等于端点，否则在前两个判断中）
           'For j% = 2 To m_lin(t_line%).data(0).data0.in_point(0) - 1
           If compare_two_point(m_poi(m_lin(t_line%).data(0).data0.in_point(t_p% - 1)).data(0).data0.coordinate, _
                      m_poi(m_lin(t_line%).data(0).data0.in_point(t_p%)).data(0).data0.coordinate, _
                       m_lin(t_line%).data(0).data0.poi(0), _
                        m_lin(t_line%).data(0).data0.poi(1), 4) = -1 Then  '移动点与前一点顺序变化
                    Call exchange_two_integer(m_lin(t_line%).data(0).data0.in_point(t_p% - 1), _
                                         m_lin(t_line%).data(0).data0.in_point(t_p%)) '交换两点
                If t_p% = 2 Then '移动点是第二点，与前一点交换机后，改变了前端点的坐标
                        m_lin(t_line%).data(0).data0.end_point_coord(0) = m_poi(point_no%).data(0).data0.coordinate
                            m_lin(t_line%).data(0).is_change = 255
                             'Call C_display_picture.re_draw_line(t_line%)
                End If
           ElseIf compare_two_point(m_poi(m_lin(t_line%).data(0).data0.in_point(t_p%)).data(0).data0.coordinate, _
                       m_poi(m_lin(t_line%).data(0).data0.in_point(t_p% + 1)).data(0).data0.coordinate, _
                        m_lin(t_line%).data(0).data0.poi(0), _
                         m_lin(t_line%).data(0).data0.poi(1), 4) = -1 Then
                    Call exchange_two_integer(m_lin(t_line%).data(0).data0.in_point(t_p%), _
                                         m_lin(t_line%).data(0).data0.in_point(t_p% + 1))
                If t_p% = m_lin(t_line%).data(0).data0.in_point(0) - 1 Then '移动点是倒数第二点，与后一点交换机后，改变了后端点的坐标
                            m_lin(t_line%).data(0).data0.end_point_coord(1) = m_poi(point_no%).data(0).data0.coordinate
                            m_lin(t_line%).data(0).is_change = 255
                            'Call C_display_picture.re_draw_line(t_line%)
                End If
           End If
          End If
Else
End If
   If pn% = m_lin(t_line%).data(0).data0.in_point(1) And _
       p1% = m_lin(t_line%).data(0).data0.in_point(m_lin(t_line%).data(0).data0.in_point(0)) Then
        For i% = 1 To 9
         temp_color(i%) = m_lin(t_line%).data(0).data0.color(i%)
        Next i%
        For i% = 1 To m_lin(t_line%).data(0).data0.in_point(0)
         m_lin(t_line%).data(0).data0.color(i%) = _
            temp_color(m_lin(t_line%).data(0).data0.in_point(0) - i% + 1)
        Next i%
   End If
 Call add_next_change_element(line_, t_line%)
'Call C_display_picture.re_draw_line(t_line%)
End Sub

Public Sub change_ratio_for_measure()
Dim i%
    For i% = 1 To Ratio_for_measure.sons.last_son
     If Ratio_for_measure.sons.son(i%).ty = line_ Then
     ElseIf Ratio_for_measure.sons.son(i%).ty = circle_ Then
      Call change_m_circle(Ratio_for_measure.sons.son(i%).no, depend_condition(Ratio_for_measure_, 0))
     ElseIf Ratio_for_measure.sons.son(i%).ty = wenti_cond_ Then
     End If
 Next i%

End Sub


Public Sub change_epolygon(epolygon_no%)
'无图
Dim i%, l%, j%, n%
Dim A!
Dim t_coord As POINTAPI
n% = epolygon(epolygon_no%).data(0).p.total_v
A! = PI * (n% - 2) / n%
   epolygon(epolygon_no%).data(0).p.coord(0) = m_poi(epolygon(epolygon_no%).data(0).p.v(0)).data(0).data0.coordinate
   epolygon(epolygon_no%).data(0).p.coord(1) = m_poi(epolygon(epolygon_no%).data(0).p.v(1)).data(0).data0.coordinate
For i% = 2 To n% - 1
  If epolygon(epolygon_no%).data(0).p.direction = -1 Then
   epolygon(epolygon_no%).data(0).p.coord(i%).X = epolygon(epolygon_no%).data(0).p.coord(i% - 1).X + _
    (epolygon(epolygon_no%).data(0).p.coord(i% - 2).X - _
       epolygon(epolygon_no%).data(0).p.coord(i% - 1).X) * Cos(A!) - _
        (epolygon(epolygon_no%).data(0).p.coord(i% - 2).Y - _
         epolygon(epolygon_no%).data(0).p.coord(i% - 1).Y) * Sin(A!)
   epolygon(epolygon_no%).data(0).p.coord(i%).Y = epolygon(epolygon_no%).data(0).p.coord(i% - 1).Y + _
    (epolygon(epolygon_no%).data(0).p.coord(i% - 2).X - _
      epolygon(epolygon_no%).data(0).p.coord(i% - 1).X) * Sin(A!) + _
     (epolygon(epolygon_no%).data(0).p.coord(i% - 2).Y - _
       epolygon(epolygon_no%).data(0).p.coord(i% - 1).Y) * Cos(A!)
  Else
   epolygon(epolygon_no%).data(0).p.coord(i%).X = epolygon(epolygon_no%).data(0).p.coord(i% - 1).X + _
    (epolygon(epolygon_no%).data(0).p.coord(i% - 2).X - _
     epolygon(epolygon_no%).data(0).p.coord(i% - 1).X) * Cos(A!) + _
     (epolygon(epolygon_no%).data(0).p.coord(i% - 2).Y - _
       epolygon(epolygon_no%).data(0).p.coord(i% - 1).Y) * Sin(A!)
   epolygon(epolygon_no%).data(0).p.coord(i%).Y = epolygon(epolygon_no%).data(0).p.coord(i% - 1).Y - _
   (epolygon(epolygon_no%).data(0).p.coord(i% - 2).X - _
     epolygon(epolygon_no%).data(0).p.coord(i% - 1).X) * Sin(A!) + _
     (epolygon(epolygon_no%).data(0).p.coord(i% - 2).Y - _
       epolygon(epolygon_no%).data(0).p.coord(i% - 1).Y) * Cos(A!)
  End If
   m_poi(epolygon(epolygon_no%).data(0).p.v(i%)).data(0).data0.coordinate = epolygon(epolygon_no%).data(0).p.coord(i%)
     Call change_m_point(epolygon(epolygon_no%).data(0).p.v(i%))
Next i%
End Sub

Public Sub add_next_change_element(ty As Integer, no As Integer)
Dim temp_no%
If ty = point_ Then
 m_poi(no).data(0).is_change = False
 If m_poi(no).data(0).data0.visible = 0 Then
    Exit Sub
 End If
 temp_no% = change_picture_start_no.change_point_start_no
 If temp_no% = 0 Then
    change_picture_start_no.change_point_start_no = no
    temp_no% = no
    change_picture_start_no.current_point_no = no
    m_poi(temp_no%).data(0).next_no = 0
    change_picture_start_no.is_picture_change = True
 End If
 Do While temp_no% > 0
   If temp_no% = no Then
    Exit Sub
   End If
   temp_no% = m_poi(temp_no%).data(0).next_no
 Loop
   m_poi(change_picture_start_no.current_point_no).data(0).next_no = no%
   change_picture_start_no.current_point_no = no
   m_poi(no).data(0).next_no = 0
 '***********************************************************************
ElseIf ty = line_ Then
 m_lin(no).data(0).is_change = False
 If m_lin(no).data(0).data0.visible = 0 Then
    Exit Sub
 End If
 temp_no% = change_picture_start_no.change_line_start_no
 If temp_no% = 0 Then
    change_picture_start_no.change_line_start_no = no
    temp_no% = no
    change_picture_start_no.current_line_no = no
    m_lin(temp_no%).data(0).next_no = 0
    change_picture_start_no.is_picture_change = True
 End If
 Do While temp_no% > 0
   If temp_no% = no Then
    Exit Sub
   End If
      temp_no% = m_lin(temp_no%).data(0).next_no
 Loop
   m_lin(change_picture_start_no.current_line_no).data(0).next_no = no
   change_picture_start_no.current_line_no = no
   m_lin(no).data(0).next_no = 0
'**********************************************************************
ElseIf ty = circle_ Then
 m_Circ(no).data(0).is_change = False
 If m_Circ(no).data(0).data0.visible = 0 Then
    Exit Sub
 End If
 temp_no% = change_picture_start_no.change_circle_start_no
 If temp_no% = 0 Then
    change_picture_start_no.change_circle_start_no = no
    temp_no% = no
    change_picture_start_no.current_circle_no = no
    m_Circ(temp_no%).data(0).next_no = 0
    change_picture_start_no.is_picture_change = True
 End If
 Do While temp_no% > 0
   If temp_no% = no Then
    Exit Sub
   End If
   temp_no% = m_Circ(temp_no%).data(0).next_no
 Loop
   m_Circ(change_picture_start_no.current_circle_no).data(0).next_no = no
   change_picture_start_no.current_circle_no = no
   m_Circ(no).data(0).next_no = 0
End If
End Sub

Public Sub draw_change_picture()
Dim temp_no%
temp_no% = change_picture_start_no.change_point_start_no
Do While temp_no% > 0
   Call C_display_picture.draw_point(temp_no%)
   temp_no% = m_poi(temp_no%).data(0).next_no
Loop
change_picture_start_no.change_point_start_no = 0
'***********************************************************
temp_no% = change_picture_start_no.change_line_start_no
Do While temp_no% > 0
   Call C_display_picture.re_draw_line(temp_no%)
   temp_no% = m_lin(temp_no%).data(0).next_no
Loop
change_picture_start_no.change_line_start_no = 0
'************************************************************
temp_no% = change_picture_start_no.change_circle_start_no
Do While temp_no% > 0
   Call C_display_picture.draw_circle(temp_no%, 0, 0)
   temp_no% = m_Circ(temp_no%).data(0).next_no
Loop
change_picture_start_no.change_circle_start_no = 0
change_picture_start_no.is_picture_change = False

End Sub

Attribute VB_Name = "DRAW"
Option Explicit
Global yidian_no As Integer
Global yidian_step% '　移点
Global input_point_type%
Global draw_line_no As Integer
Dim draw_circle_no As Integer
Dim ele1 As condition_type
Dim ele2 As condition_type
Dim init_p As POINTAPI
Dim temp_r!
Dim change_fig_type As Byte
Dim yidian_statue As Boolean
Dim move_direct As Byte
Dim Up_Enabled As Boolean 'mouse_up 作用
Dim Move_Enabled As Boolean 'mouse_move 作用=TRUE 可以作移动
Dim Move_statue As Byte '0 未有移动1左键移动 2右键移动
 Const red = True
 Const blue = False
Dim move_init As Integer
Type control_data_type
select_point(5) As Integer
select_line(5) As Integer
select_circle(5) As Integer
select_plane(5) As Integer
forbid_point(5) As Integer
forbid_line(5) As Integer
forbid_circle(5) As Integer
forbid_plane(5) As Integer
End Type
Dim control_data   As control_data_type
Dim mouse_move_coord As POINTAPI
Dim mouse_up_coord As POINTAPI
Dim mouse_down_coord As POINTAPI
Dim t_same_points(8) As Integer '记录两个序列的相同点
Dim temp_circles_for_draw(0 To 7) As Integer
Dim temp_lines_for_draw(0 To 7) As Integer
'Global input_coord '  '鼠标输入的坐标
Global time_no  As Integer   'draw_form.timer 　的状态记录

Public Sub draw_red_line(l%)

If m_lin(l%).data(0).data0.in_point(0) > 0 Then
If m_lin(l%).data(0).data0.visible = 2 Then
Draw_form.Line (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.X, _
    m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.Y)- _
     (m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.X, _
       m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.Y), QBColor(condition_color)
End If
Draw_form.Line (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.X, _
    m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.Y)- _
     (m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.X, _
       m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.Y), QBColor(12)
Else
If m_lin(m_lin(l%).data(0).data0.in_point(1)).data(0).data0.visible = 2 Then
Draw_form.Line (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.X, _
    m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.Y)- _
     (m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.X, _
       m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.Y), QBColor(condition_color)
End If
Draw_form.Line (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.X, _
    m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.Y)- _
     (m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.X, _
       m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.Y), QBColor(12)
End If

 
   Call set_red_line(l%)
End Sub

Private Sub draw_aid_line(input_coord As POINTAPI, ty As Integer) '画辅助图 ty=1 仅有一个端点的辅助线，_
                     ' 如画切先，ty=0 两个端点的直线延长线
Dim ele1 As condition_type
Dim ele2 As condition_type
Dim i%, tp%
Dim out_coord As POINTAPI
input_point_type% = read_inter_point(input_coord, ele1, ele2, tp%, False, 0)  '读取点;ty=0 不读单点线
         If input_point_type% = exist_point Then '读已知取点
           i% = ele1.no '已知点所在的线或圆
         End If
If operator = "ask" And ty = 0 Then '提问 LoadResString_(140)
     If tp% > 0 Then '显示点的信息
        Draw_form.Label2.Caption = m_poi(tp%).data(0).inform
     Else
      If ele1.no > 0 And ele2.no > 0 Then
         If ele1.ty = line_ And ele2.ty = line_ Then
           Draw_form.Label2.Caption = LoadResString_(1780, _
                          "\\1\\" + set_display_line(ele1.no) + _
                          "\\2\\" + set_display_line(ele2.no))
         ElseIf ele1.ty = line_ And ele2.ty = circle_ Then
           Draw_form.Label2.Caption = LoadResString_(1780, _
                         "\\1\\" + set_display_line(ele1.no) + _
                         "\\2\\" + set_display_circle(ele2.no))
         ElseIf ele1.ty = circle_ And ele2.ty = circle_ Then
           Draw_form.Label2.Caption = LoadResString_(1780, _
                         "\\1\\" + set_display_circle(ele1.no) + _
                         "\\2\\" + set_display_circle(ele2.no))
         End If
      ElseIf ele1.no > 0 Then
        If ele1.ty = point_ Then
           Draw_form.Label2.Caption = m_poi(ele1.no).data(0).inform
        ElseIf ele1.ty = line_ Then
           Draw_form.Label2.Caption = m_lin(ele1.no).data(0).inform '显示直线的信息
        ElseIf ele1.ty = circle_ Then
           Draw_form.Label2.Caption = m_Circ(ele1.no).data(0).inform '显示圆的信息
        End If
      End If
    End If
    '********************************************************************************
    '设置Draw_form.Label2.Caption的位置
    If Draw_form.Label2.Caption <> "" Then
      If Int(input_coord.X) < Draw_form.Label2.width Then
        Draw_form.Label2.left = input_coord.X + 4
      Else
        Draw_form.Label2.left = input_coord.X - Draw_form.Label2.width - 4
      End If
      If Int(input_coord.Y) < 50 Then
         Draw_form.Label2.top = input_coord.Y + 4
      Else
           Draw_form.Label2.top = input_coord.Y - Draw_form.Label2.Height - 4
      End If
      Draw_form.Label2.visible = True '显示信息
    End If
Else 'If ele1.ty = line_ Or ele2.ty = line_ Then
'*******************************************************************************************
'     Call draw_temp_line_for_move(point_type%, ele1.no, ele2.no)
       If ele1.ty <> line_ Then
          ele1.no = 0
       End If
       If ele2.ty <> line_ Then
          ele2.no = 0
       End If
       If ele1.no > 0 Or ele2.no > 0 Then
         Call C_display_picture.draw_aid_line(ele1.no, ele2.no, input_coord.X, input_coord.Y)
       End If
End If
End Sub

Public Sub redraw_red_line(l%)
Dim i%
For i% = 0 To 15
 If red_line(i%) = l% Then

If m_lin(l%).data(0).data0.in_point(0) > 0 Then
Draw_form.Line (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.X, _
    m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.Y)- _
     (m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.X, _
       m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.Y), QBColor(12)
 If m_lin(l%).data(0).data0.visible = 2 Then
Draw_form.Line (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.X, _
    m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.Y)- _
     (m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.X, _
       m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.Y), _
         QBColor(condition_color)
 End If
Else
Draw_form.Line (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.X, _
    m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.Y)- _
     (m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.X, _
       m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.Y), QBColor(12)
 If m_lin(m_lin(l%).data(0).data0.in_point(1)).data(0).data0.visible = 2 Then
Draw_form.Line (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.X, _
    m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.Y)- _
     (m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.X, _
       m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.Y), _
         QBColor(condition_color)
 End If

End If
  red_line(i%) = 0
End If
 Next i%
End Sub

Public Sub draw_grey_line(l%)
Draw_form.Line (m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.X, _
    m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.Y)- _
     (m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.X, _
       m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.Y), _
         QBColor(fill_color)
   Call set_grey_line(l%)
End Sub
Public Sub draw_arc(ob As Object, ByVal cx As Single, ByVal cy As Single, ByVal startx As Single, _
       ByVal starty As Single, ByVal endx As Single, ByVal endy As Single, ByVal color%)
Dim r As Single
'Dim fm As Object
'Set fm = Form1
    r = sqr((startx - cx) ^ 2 + (starty - cy) ^ 2)
If starty < cy And startx > cx And endx > cx And endy < cy Then '第一象限到第一象限
         ob.Circle (cx, cy), r, QBColor(color%), Atn((cy - starty) / (startx - cx)), Atn((cy - endy) / (endx - cx))
ElseIf startx > cx And starty < cy And endx < cx And endy < cy Then '第一象限到第二象限
         ob.Circle (cx, cy), r, QBColor(color%), Atn((cy - starty) / (startx - cx)), PI / 2 + Atn((cy - endy) / (cx - endx))
ElseIf startx > cx And starty < cy And endx < cx And endy > cy Then '第一象限到第三象限
         ob.Circle (cx, cy), r, QBColor(color%), Atn((cy - starty) / (startx - cx)), PI + Atn((endy - cy) / (cx - endx))
ElseIf startx > cx And starty < cy And endx > cx And endy > cy Then '第一象限到第四象限
         ob.Circle (cx, cy), r, QBColor(color%), Atn((cy - starty) / (startx - cx)), 2 * PI - Atn((endy - cy) / (endx - cx))
ElseIf startx < cx And starty < cy And endx > cx And endy < cy Then '第二象限到第一象限
         ob.Circle (cx, cy), r, QBColor(color%), PI / 2 + Atn((cy - starty) / (cx - startx)), Atn((cy - endy) / (endx - cx))
ElseIf startx < cx And starty < cy And endx < cx And endy < cy Then '第二象限到第二象限
         ob.Circle (cx, cy), r, QBColor(color%), PI / 2 + Atn((cy - starty) / (cx - startx)), PI / 2 + Atn((cy - endy) / (cx - endx))
ElseIf startx < cx And starty < cy And endx < cx And endy > cy Then '第二象限到第三象限
         ob.Circle (cx, cy), r, QBColor(color%), PI / 2 + Atn((cy - starty) / (cx - startx)), PI + Atn((endy - cy) / (cx - endx))
ElseIf startx < cx And starty < cy And endx > cx And endy > cy Then '第二象限到第四象限
         ob.Circle (cx, cy), r, QBColor(color%), PI / 2 + Atn((cy - starty) / (cx - startx)), 2 * PI - Atn((endy - cy) / (endx - cx))
ElseIf starty > cy And startx < cx And endx > cx And endy < cy Then '第三象限到第一象限
         ob.Circle (cx, cy), r, QBColor(color%), PI + Atn((starty - cy) / (cx - startx)), Atn((cy - endy) / (endx - cx))
ElseIf startx < cx And starty > cy And endx < cx And endy < cy Then '第三象限到第二象限
         ob.Circle (cx, cy), r, QBColor(color%), PI + Atn((starty - cy) / (cx - startx)), PI / 2 + Atn((endy - cy) / (endx - cx))
ElseIf startx < cx And starty > cy And endx < cx And endy > cy Then '第三象限到第三象限
         ob.Circle (cx, cy), r, QBColor(color%), PI + Atn((starty - cy) / (cx - startx)), PI + Atn((endy - cy) / (cx - endx))
ElseIf startx < cx And starty > cy And endx > cx And endy > cy Then '第三象限到第四象限
         ob.Circle (cx, cy), r, QBColor(color%), PI + Atn((starty - cy) / (cx - startx)), 2 * PI - Atn((endy - cy) / (endx - cx))
ElseIf starty > cy And startx > cx And endy < cy And endx > cx Then '第四象限到第一象限
         ob.Circle (cx, cy), r, QBColor(color%), 2 * PI - Atn((starty - cy) / (startx - cx)), Atn((cy - endy) / (endx - cx))
ElseIf startx > cx And starty > cy And endx < cx And endy < cy Then '第四象限到第二象限
         ob.Circle (cx, cy), r, QBColor(color%), 2 * PI - Atn((starty - cy) / (startx - cx)), PI / 2 + Atn((cy - endy) / (cx - endx))
ElseIf startx > cx And starty > cy And endx < cx And endy > cy Then '第四象限到第三象限
         ob.Circle (cx, cy), r, QBColor(color%), 2 * PI - Atn((starty - cy) / (startx - cx)), PI + Atn((endy - cy) / (cx - endx))
ElseIf startx > cx And starty > cy And endx > cx And endy > cy Then '第四象限到第四象限
         ob.Circle (cx, cy), r, QBColor(color%), 2 * PI - Atn((starty - cy) / (startx - cx)), 2 * PI - Atn((endy - cy) / (endx - cx))
End If
Exit Sub
End Sub





Public Sub remove_uncomplete_operat(op As String)
Dim i%, j%
If input_text_statue Then
MDIForm1.Text1.visible = False
MDIForm1.Text2.visible = False
 input_text_statue = False
End If
Draw_form.Picture1.visible = False
Call draw_ruler(Ratio_for_measure.Ratio_for_measure, delete)
     Draw_form.HScroll1.visible = False
For i% = 0 To 15
If red_line(i%) > 0 Then
red_line(i%) = 0
End If
Next i%
list_type_for_draw = 0
Call init_draw_data
yidian_type = 0
measur_step = 0
'last_conditions.last_cond(1).last_view_point_no = 0
If op <> "change_picture" Then
 MDIForm1.set_picture_for_change.Enabled = True
 MDIForm1.set_change_type.Enabled = False
 last_conditions.last_cond(1).change_picture_type = 0
End If
If op = "move_point" Then
ElseIf op = "measure" Then
ElseIf op = "" Then
ElseIf op = "" Then

End If
If yidian_type = 25 Then
 'Call C_display_picture.m_BPset(Draw_form, line_for_move.coord(0).x, _
         line_for_move.coord(0).y, "", display)
 'Call C_display_picture.m_BPset(Draw_form, line_for_move.coord(1).x, _
         line_for_move.coord(1).y, "", display)
 Draw_form.Line (line_for_move.coord(0).X, line_for_move.coord(0).Y)- _
      (line_for_move.coord(1).X, line_for_move.coord(1).Y), QBColor(fill_color)
ElseIf yidian_type = 26 Then
 C_curve.Class_Init
 'Draw_form.Circle (circle_for_move.center_coord.X, circle_for_move.center_coord.Y), _
        circle_for_move.radii, QBColor(fill_color)
End If
If set_change_fig > 0 And is_first_move Then
 If last_conditions.last_cond(1).change_picture_type = line_ Then
  Call C_display_picture.draw_red_point(line_for_change.line_no(0).poi(0))
  Call C_display_picture.draw_red_point(line_for_change.line_no(0).poi(1))
  'Call draw_red_line(line_number5(line_for_change.poi(0), _
     line_for_change.poi(1), 0, 0, 0))
ElseIf last_conditions.last_cond(1).change_picture_type = polygon_ Then
 For i% = 1 To Polygon_for_change.p(0).total_v - 1
 Call C_display_picture.draw_red_point(Polygon_for_change.p(0).v(i%))
 Call draw_red_line(line_number0(Polygon_for_change.p(0).v(i%), _
     Polygon_for_change.p(0).v(i% - 1), 0, 0))
 Next i%
 Call C_display_picture.draw_red_point(Polygon_for_change.p(0).v(0))
 If Polygon_for_change.p(0).total_v > 2 Then
 Call draw_red_line(line_number0(Polygon_for_change.p(0).v(0), _
     Polygon_for_change.p(0).v(Polygon_for_change.p(0).total_v - 1), 0, 0))
 End If
 ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
 'Draw_form.Circle (circle_for_move.center_coord.X, circle_for_move.center_coord.Y), _
        circle_for_move.radii, QBColor(fill_color)
 End If
End If
Call recove_set_menu_for_set_function_data
'MDIForm1.StatusBar1.Panels(1).text = ""
End Sub

Sub centroid(ByVal p1%, ByVal p2%, ByVal p3%, p%, p4%, p5%, p6%, _
              is_draw As Boolean) '画重心 ***12.30
Dim X%, Y%, t_p1%, t_p2%, t_p3%, l%, m%, n%
Call draw_triangle(p1%, p2%, p3%, condition)
   Call mid_point(p2%, p3%, p4%, is_draw)
    Call mid_point(p1%, p3%, p5%, is_draw)
     Call mid_point(p1%, p2%, p6%, is_draw)
      l% = line_number(p1%, p4%, pointapi0, pointapi0, _
                       depend_condition(point_, p1%), depend_condition(point_, p2%), _
                       condition, condition_color, 1, 0)
      m% = line_number(p2%, p5%, pointapi0, pointapi0, _
                       depend_condition(point_, p2%), depend_condition(point_, p5%), _
                       condition, condition_color, 1, 0)
      n% = line_number(p3%, p6%, pointapi0, pointapi0, _
                       depend_condition(point_, p3%), depend_condition(point_, p6%), _
                       condition, condition_color, 1, 0)
      Call set_Divide_Point(p1%, p4%, 2, p%, is_draw)
      record_0.data0.condition_data.condition_no = 0
       Call add_point_to_line(p%, m%, 0, False, False, 0)
        Call add_point_to_line(p%, n%, 0, False, False, 0)
End Sub

Sub centroid1(ByVal p1%, ByVal p2%, ByVal p3%, p%, p4%, p5%, p6%, is_change As Boolean) '重心 用于重画 12.30


Dim X%, Y%, t_p1%, t_p2%, t_p3%, l%, m%, n%
   Call mid_point1(p2%, p3%, p4%, is_change)
    Call mid_point1(p1%, p3%, p5%, is_change)
     Call mid_point1(p1%, p2%, p6%, is_change)
      Call ratio_point1(p1%, p4%, 2, 1, p%, is_change)


End Sub
Function compare_two_point(p_coord1 As POINTAPI, p_coord2 As POINTAPI, ByVal p3%, ByVal p4%, interval%) As Integer    '两点在平面上排序12.30
'1 升序
Dim r, k1 As Single
Dim tp(3) As Integer
Dim dr(1) As Integer
Dim d(3) As Long
Dim k%
  tp(2) = Abs(p3%) '点的序号
   tp(3) = Abs(p4%)
    compare_two_point = 0
d(0) = Abs(p_coord1.X - p_coord2.X) '两点的距离
d(1) = Abs(p_coord1.Y - p_coord2.Y)
If tp(2) > 0 And tp(3) > 0 Then '
d(2) = Abs(m_poi(tp(2)).data(0).data0.coordinate.X - m_poi(tp(3)).data(0).data0.coordinate.X) '后两点的距离
d(3) = Abs(m_poi(tp(2)).data(0).data0.coordinate.Y - m_poi(tp(3)).data(0).data0.coordinate.Y)
End If
If d(0) > d(1) Then '前两点横坐标的差大于纵坐标的差
   dr(0) = 0
Else
   dr(0) = 1
   Call exchange_two_long_integer(d(0), d(1)) '
End If
If tp(2) > 0 And tp(3) > 0 Then
If d(2) > d(3) Then '后两点横坐标的差大于纵坐标的差
   dr(1) = 0
Else
   dr(1) = 1
   Call exchange_two_long_integer(d(2), d(3))
End If
If d(0) < d(2) Then
  dr(0) = dr(1)
End If
End If
If tp(2) = 0 Or tp(3) = 0 Then
 r = CSng((p_coord1.X - p_coord2.X)) ^ 2 + _
     CSng((p_coord1.Y - p_coord2.Y)) ^ 2
  If r < 10 Then  '         '  两点距离小于3.2
   compare_two_point = 0 '两点太近，不能排序
    Exit Function
  Else
   If dr(0) = 0 Then ''前两点横坐标的差大于纵坐标的差
    compare_two_point = Sgn(p_coord2.X - p_coord1.X)
   Else
    compare_two_point = Sgn(p_coord1.Y - p_coord2.Y)
   End If
  End If
Else '有后两点做标准
  If dr(0) = 0 Then ''前两点横坐标的差大于纵坐标的差
    compare_two_point = Sgn((p_coord1.X - p_coord2.X)) * _
                  Sgn((m_poi(tp(2)).data(0).data0.coordinate.X - m_poi(tp(3)).data(0).data0.coordinate.X)) '派序点与标准点同号=1,否则=-1
  Else
    compare_two_point = Sgn((p_coord1.Y - p_coord2.Y)) * _
                  Sgn((m_poi(tp(2)).data(0).data0.coordinate.Y - m_poi(tp(3)).data(0).data0.coordinate.Y))
  End If
End If
End Function
Public Sub change_ratio()
Dim n As Byte
For n = 1 To last_length
Call display_m_string(length_(n).string_no, no_display)

Call length_of_line(length_(n))
   Measur_string(length_(n).string_no) = LoadResString_(1615, _
          "\\1\\" + m_poi(length_(n).poi(0)).data(0).data0.name + _
                 m_poi(length_(n).poi(1)).data(0).data0.name + _
          "\\2\\" + str(length_(n).len / Ratio_for_measure.Ratio_for_measure))
Call display_m_string(length_(n).string_no, display)
            
Next n
End Sub
Public Sub display_m_string(n As Byte, display_or_delete As Boolean)
   Wenti_form.Picture3.CurrentY = 70 + 20 * n
   Wenti_form.Picture3.CurrentX = 20
   If display_or_delete = 0 Then
   Call SetTextColor(Wenti_form.Picture3.hdc, QBColor(15))
   Else
   Call SetTextColor(Wenti_form.Picture3.hdc, QBColor(9))
   End If
   Wenti_form.Picture3.FontSize = 10
   Wenti_form.Picture3.Print Measur_string(n);
    Call SetTextColor(Wenti_form.Picture3.hdc, QBColor(9))
    Draw_form.FontSize = 8.25
  ' If n = 0 Then
   '     Text1.Top = CurrentY - 2
    '     Text1.Left = CurrentX
   'End If
End Sub



Sub draw_plus_point(ob As Object, ByVal p%, coord As POINTAPI, _
     display_or_delete As Boolean)  ', col%,display_or_delete as boolean )  '???????
  Dim n$
  If p% > 0 Then
   n$ = m_poi(p%).data(0).data0.name
   If line_width < 3 Then
   ob.PaintPicture Draw_form.ImageList2.ListImages(Asc(n$) - 64).Picture, coord.X - 2, _
              coord.Y - 2, 16, 18, 0, 0, 16, 18, &H990066
   Else
   ob.PaintPicture Draw_form.ImageList4.ListImages(Asc(n$) - 64).Picture, coord.X - 5, _
              coord.Y - 5, 32, 32, 0, 0, 32, 32, &H990066
   End If
  End If
End Sub



Sub draw_three_point_circle(x1&, y1&, x2&, y2&, X3&, Y3&, _
            c As Long, c_x&, c_y&, r&, dis_or_no As Boolean, ob As Byte) 'c% color
Dim temp_num0&, temp_num1&, temp_num2&
        temp_num0& = x1& * y2& + x2& * Y3& + _
           X3& * y1& - x1& * Y3& - x2& * y1& - X3& * y2& _
             'is_in_line(poi1%, poi2%, i%)  '(two_point_of_line(1)).data(0).data0.coordinate.X * poi(two_point_of_line(0)).data(0).data0.coordinate.Y - poi(two_point_of_line(0)).data(0).data0.coordinate.X * poi(two_point_of_line(1)).data(0).data0.coordinate.Y + poi(i%).data(0).data0.coordinate.X * poi(two_point_of_line(1)).data(0).data0.coordinate.Y - poi(two_point_of_line(1)).data(0).data0.coordinate.X * poi(i%).data(0).data0.coordinate.Y
         If temp_num0& = 0 Then
          If Abs(x1& - X3&) > 4 Then
            If (x1& - x2&) * (x2& - X3&) > 0 Then
             If ob = 0 Then
             Draw_form.Line (x1&, y1&)-(X3&, Y3&), c
             Else
             Draw_form.Line (x1&, y1&)-(X3&, Y3&), c
             End If
            ElseIf (x2& - x1&) * (x1& - X3&) > 0 Then
             If ob = 0 Then
             Draw_form.Line (x2&, y2&)-(X3&, Y3&), c
             Else
             Draw_form.Line (x2&, y2&)-(X3&, Y3&), c
             End If
            ElseIf (x1& - X3&) * (X3& - x2&) > 0 Then
             If ob = 0 Then
             Draw_form.Line (x1&, y1&)-(x2&, y2&), c
             Else
             Draw_form.Line (x1&, y1&)-(x2&, y2&), c
             End If
            End If
          Else
             If (y1& - y2&) * (y2& - Y3&) > 0 Then
              If ob = 0 Then
              Draw_form.Line (x1&, y1&)-(X3&, Y3&), c
              Else
              Draw_form.Line (x1&, y1&)-(X3&, Y3&), c
              End If
             ElseIf (y2& - y1&) * (y1& - Y3&) > 0 Then
              If ob = 0 Then
              Draw_form.Line (x2&, y2&)-(X3&, Y3&), c
              Else
              Draw_form.Line (x2&, y2&)-(X3&, Y3&), c
              End If
             ElseIf (y1& - Y3&) * (Y3& - y2&) > 0 Then
              If ob = 0 Then
              Draw_form.Line (x1&, y1&)-(x2&, y2&), c
              Else
              Draw_form.Line (x1&, y1&)-(x2&, y2&), c
              End If
             End If
          End If
            Exit Sub
Else

           temp_num1& = (y1& - y2&) * (x2& ^ 2 - X3& ^ 2) - _
              (y2& - Y3&) * (x1& ^ 2 - x2& ^ 2)
            temp_num2& = (y1& - y2&) * (y2& - Y3&) * (Y3& - y1&)
             c_x& = -(temp_num1& + temp_num2&) / temp_num0& / 2
        temp_num1& = (x1& - x2&) * (y2& ^ 2 - Y3& ^ 2) - _
          (x2& - X3&) * (y1& ^ 2 - y2& ^ 2)
         temp_num2& = (x1& - x2&) * (x2& - X3&) * (X3& - x1&)
          c_y& = (temp_num1& + temp_num2&) / temp_num0& / 2
          r& = sqr((x1& - c_x&) ^ 2 + (y1& - c_y&) ^ 2)
        If dis_or_no = display Then
         If ob = 0 Then
         Draw_form.Circle (c_x&, c_y&), r&, c
         Else
         Draw_form.Circle (c_x&, c_y&), r&, c
         End If
        End If
         End If

End Sub
Public Sub re_name_point(ByVal p%)
   choose_point = p%
    If choose_point > 0 Then
     re_name_ty = 0
     Call C_display_picture.draw_red_point(choose_point)
      yidian_stop = False
      Call C_display_picture.flash_point(choose_point)
       Draw_form.SetFocus
        operator = "re_name"
   End If
End Sub
Sub incenter(p1%, p2%, p3%, p4%, cr%, t_p1%, _
              t_p2%, t_p3%, is_change As Boolean) '内心
Dim A!, b!, c!, d!, p!, q!, w!, r!, t_circle%, i%
Dim X%, Y%
A! = sqr((m_poi(p2%).data(0).data0.coordinate.X - m_poi(p3%).data(0).data0.coordinate.X) ^ 2 + (m_poi(p2%).data(0).data0.coordinate.Y - m_poi(p3%).data(0).data0.coordinate.Y) ^ 2)
 b! = sqr((m_poi(p3%).data(0).data0.coordinate.X - m_poi(p1%).data(0).data0.coordinate.X) ^ 2 + (m_poi(p3%).data(0).data0.coordinate.Y - m_poi(p1%).data(0).data0.coordinate.Y) ^ 2)
  c! = sqr((m_poi(p1%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) ^ 2 + (m_poi(p1%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y) ^ 2)
p! = (A! + b! + c!) / 2
 d! = p! - A!
  q! = p! - b!
w! = p! - c!
r! = sqr(q! * d! * w! / p!) '/p!
q! = 1 / A! + 1 / b! + 1 / c!
t_coord.X = Int((m_poi(p1%).data(0).data0.coordinate.X * A! + m_poi(p2%).data(0).data0.coordinate.X * b! + m_poi(p3%).data(0).data0.coordinate.X * c!) / p! / 2)
t_coord.Y = Int((m_poi(p1%).data(0).data0.coordinate.Y * A! + m_poi(p2%).data(0).data0.coordinate.Y * b! + m_poi(p3%).data(0).data0.coordinate.Y * c!) / p! / 2)
Call draw_triangle(p1%, p2%, p3%, condition)
  If p4% = 0 Then
  last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1 '***
  'MDIForm1.Toolbar1.Buttons(21).Image = 33
   p4% = last_conditions.last_cond(1).point_no '***
  End If
  Call set_point_coordinate(p4%, t_coord, is_change)
  'poi(p4%).data(0).data0.coordinate.Y = Y%
  Call set_point_visible(p4%, 1, False)
  m_poi(p4%).data(0).degree = 0
  'Call set_point_color(p4%, 0)
'Call draw_circle(Draw_form, Circ(cr%).data(0).data0)
'Call draw_point(Draw_form, poi(p4%), 0, display)
Call orthofoot(p4%, p2%, p3%, t_p1%, 0, False)
Call orthofoot(p4%, p3%, p1%, t_p2%, 0, False)
Call orthofoot(p4%, p1%, p2%, t_p3%, 0, False)
Call add_point_to_m_circle(t_p1%, cr%, True)
Call add_point_to_m_circle(t_p2%, cr%, True)
Call add_point_to_m_circle(t_p3%, cr%, True)
   Call m_circle_number(1, p4%, m_poi(p4%).data(0).data0.coordinate, t_p1%, t_p2%, t_p3%, r!, 0, 0, _
             1, 1, condition, condition_color, True)
End Sub

Sub incenter1(p1%, p2%, p3%, p4%, r&, t_p1%, t_p2%, t_p3%, is_change As Boolean)
Dim A!, b!, c!, d!, p!, q!, w!, t_circle%, i%
A! = sqr((m_poi(p2%).data(0).data0.coordinate.X - m_poi(p3%).data(0).data0.coordinate.X) ^ 2 + (m_poi(p2%).data(0).data0.coordinate.Y - m_poi(p3%).data(0).data0.coordinate.Y) ^ 2)
 b! = sqr((m_poi(p3%).data(0).data0.coordinate.X - m_poi(p1%).data(0).data0.coordinate.X) ^ 2 + (m_poi(p3%).data(0).data0.coordinate.Y - m_poi(p1%).data(0).data0.coordinate.Y) ^ 2)
  c! = sqr((m_poi(p1%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) ^ 2 + (m_poi(p1%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y) ^ 2)
p! = (A! + b! + c!) / 2
 d! = p! - A!
  q! = p! - b!
w! = p! - c!
r& = sqr(q! * d! * w! / p!) '/p!
q! = 1 / A! + 1 / b! + 1 / c!
 t_coord.X = Int((m_poi(p1%).data(0).data0.coordinate.X * A! + m_poi(p2%).data(0).data0.coordinate.X * b! + m_poi(p3%).data(0).data0.coordinate.X * c!) / p! / 2)
 t_coord.Y = Int((m_poi(p1%).data(0).data0.coordinate.Y * A! + m_poi(p2%).data(0).data0.coordinate.Y * b! + m_poi(p3%).data(0).data0.coordinate.Y * c!) / p! / 2)
 Call set_point_coordinate(p4%, t_coord, is_change)
'poi(p4%).data(0).data0.coordinate.X = X%
 'poi(p4%).data(0).data0.coordinate.Y = Y%
  'circ(cr%).data(0).radii = r!
Call orthofoot1(m_poi(p4%).data(0).data0.coordinate, m_poi(p2%).data(0).data0.coordinate, _
                          m_poi(p3%).data(0).data0.coordinate, t_coord, t_p1%, is_change)
Call orthofoot1(m_poi(p4%).data(0).data0.coordinate, m_poi(p3%).data(0).data0.coordinate, _
                          m_poi(p1%).data(0).data0.coordinate, t_coord, t_p2%, is_change)
Call orthofoot1(m_poi(p4%).data(0).data0.coordinate, m_poi(p1%).data(0).data0.coordinate, _
                          m_poi(p2%).data(0).data0.coordinate, t_coord, t_p3%, is_change)
End Sub

Public Function inter_point_line_line(p1%, p2%, p3%, p4%, l1%, l2%, p%, _
                                    out_coord As POINTAPI, _
                                     is_set_data As Boolean, is_draw As Boolean, is_change As Boolean) As Integer '新点=true
'求两直线交点，可以是两条已有的直线，也可能是任意两点，但没有响应的连线
Dim i%, j%, X&, Y&, tp%  '12.30
Dim tl(1) As line_data_type
Dim ty(1) As Byte
Dim is_new_point As Boolean
Dim t_coord(3) As POINTAPI
 If l1% > 0 Then '是实线
    t_coord(0) = m_lin(l1%).data(0).data0.end_point_coord(0)
    t_coord(1) = m_lin(l1%).data(0).data0.end_point_coord(1)
 ElseIf p1% > 0 And p2% > 0 Then '未有连线
    t_coord(0) = m_poi(m_lin(l1%).data(0).data0.depend_poi(0)).data(0).data0.coordinate
    t_coord(1) = second_end_point_coordinate(l1%)
 Else
  Exit Function
 End If
   '****************************************
  If l2% > 0 Then
   t_coord(2) = m_poi(m_lin(l2%).data(0).data0.depend_poi(0)).data(0).data0.coordinate
   t_coord(3) = second_end_point_coordinate(l2%)
  ElseIf p3% > 0 And p4% > 0 Then
   t_coord(2) = m_poi(p3%).data(0).data0.coordinate
   t_coord(3) = m_poi(p4%).data(0).data0.coordinate
  Else
   Exit Function
  End If
'****************************************************************
p% = is_line_line_intersect(l1%, l2%, 0, 0, 0) '读出已有交点
 If p% = 0 Then
  is_new_point = calculate_line_line_intersect_point(t_coord(0), t_coord(1), _
                          t_coord(2), t_coord(3), _
                          out_coord, is_change) '计算交点坐标
 Else
   inter_point_line_line = exist_point '已有点
 End If
If is_set_data Then '建立数据
If l1% > 0 And l2% > 0 Then
    If p% = 0 Then '两直线没有交点
        p% = m_point_number(out_coord, condition, 1, condition_color, "", _
                               depend_condition(line_, l1%), _
                               depend_condition(line_, l2%), 0, True) '新点数据和显示
           Call add_point_to_line(p%, l1%, 0, display, is_draw, 0)
           Call add_point_to_line(p%, l2%, 0, display, is_draw, 0)
           'Call set_son_data(point_, p%, m_lin(l1%).data(0).sons)
           'Call set_son_data(point_, p%, m_lin(l2%).data(0).sons)
           m_poi(p%).data(0).degree = 0 '非自由点
           inter_point_line_line = interset_point_line_line '
    Else
           inter_point_line_line = exist_point '
    End If
ElseIf l1% = 0 Or l2% = 0 Then
   If is_new_point And is_set_data Then '有交点，且要求添加数据到数据库
            inter_point_line_line = interset_point_line_line '新交点
                   If l1% > last_conditions.last_cond(1).line_no Then 'l1%不在推理数据库中
               Call set_line_from_aid_line(l1%, p%, out_coord, l1%) '
            End If
            If l2% > last_conditions.last_cond(1).line_no Then '
               Call set_line_from_aid_line(l2%, p%, out_coord, l2%)
            End If
            If l1% = 0 And l2% = 0 Then
              inter_point_line_line = new_free_point
            ElseIf l1% = 0 Or l2% = 0 Then
              inter_point_line_line = new_point_on_line
            End If

   End If
End If
End If
     ' MDIForm1.Toolbar1.Buttons(21).Image = 33
       'If p% = 0 And is_set_data Then '原来无交点，并计算出交点坐标
       'ElseIf p% > 0 Then
       '  Call set_point_coordinate(p%, out_coord, is_change)
       '  inter_point_line_line = 0
       'End If
End Function

Sub mid_point(p1%, p2%, p%, is_draw As Boolean) '***
   Call set_Divide_Point(p1%, p2%, 1, p%, is_draw)
End Sub
Sub mid_point1(p1%, p2%, p%, is_change As Boolean)
'只求中点不画点p% = last_conditions.last_cond(1).point_no
t_coord.X = (m_poi(p1%).data(0).data0.coordinate.X + m_poi(p2%).data(0).data0.coordinate.X) / 2
t_coord.Y = (m_poi(p1%).data(0).data0.coordinate.Y + m_poi(p2%).data(0).data0.coordinate.Y) / 2
Call set_point_coordinate(p%, t_coord, is_change)
'last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
End Sub

Sub orthocenter(p1%, p2%, p3%, p%, t_p1%, t_p2%, t_p3%, ByVal no_reduce As Byte, _
         is_draw As Boolean) '***
         Call draw_triangle(p1%, p2%, p3%, condition)
Call orthofoot(p1%, p2%, p3%, t_p1%, no_reduce, is_draw) '垂足
 Call orthofoot(p2%, p3%, p1%, t_p2%, no_reduce, is_draw)
  Call orthofoot(p3%, p1%, p2%, t_p3%, no_reduce, is_draw)
   Call inter_point_line_line(p1%, t_p1%, p2%, t_p2%, _
            line_number0(p1%, t_p1%, 0, 0), _
             line_number0(p2%, t_p2%, 0, 0), p%, pointapi0, False, is_draw, False)    '两高的交点
       record_0.data0.condition_data.condition_no = 0
    Call add_point_to_line(p%, line_number(p3%, t_p3%, _
                                           pointapi0, pointapi0, _
                                           depend_condition(point_, p3%), depend_condition(point_, t_p3%), _
                                           condition, condition_color, 1, 0), _
                          0, display, is_draw, 0) '垂心在第三条高上

End Sub
Sub orthofoot(p1%, p2%, p3%, p%, ByVal no_reduce As Byte, is_draw As Boolean)
Dim i%, j%, X&, Y&
Dim p_coord As POINTAPI
Dim temp_record As record_type
Dim ty As Boolean
Call orthofoot1(m_poi(p1%).data(0).data0.coordinate, m_poi(p2%).data(0).data0.coordinate, _
   m_poi(p3%).data(0).data0.coordinate, p_coord, 0, Not is_draw)
      last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1 '***
      'MDIForm1.Toolbar1.Buttons(21).Image = 33
         p% = last_conditions.last_cond(1).point_no '***
         Call set_point_coordinate(p%, p_coord, False)
         If is_draw = True Then
         Call set_point_visible(p%, 1, False)
         m_poi(p%).data(0).degree = 0
         Else
         Call set_point_visible(p%, 1, True)
         End If
         'Call set_point_color(p%, 0)
'              Call draw_point(Draw_form, poi(p%), 0, display)
     j% = line_number(p2%, p3%, pointapi0, pointapi0, _
                      depend_condition(point_, p2%), depend_condition(point_, p3%), _
                      condition, condition_color, 1, 0)
     i% = line_number(p1%, p%, pointapi0, pointapi0, _
                      depend_condition(point_, p1%), depend_condition(point_, p%), _
                      condition, condition_color, 1, 0)
     record_0.data0.condition_data.condition_no = 0
      Call add_point_to_line(p%, j%, 0, display, is_draw, 0)
 '     ty = set_dverti(i%, j%, temp_record, 0, no_reduce)
       Call vertical_line(i%, j%, True, True)

End Sub

Sub orthofoot1(p1 As POINTAPI, p2 As POINTAPI, _
                     p3 As POINTAPI, out_p As POINTAPI, out_point_no%, Optional is_change As Boolean = False)
Dim t&, s&
'On Error GoTo orthofoot1_error
s& = (p3.X - p2.X) ^ 2 + (p3.Y - p2.Y) ^ 2
If s& = 0 Then
 out_p = p1
Else
 t& = (p3.X - p2.X) * (p2.Y - p1.Y) - _
       (p3.Y - p2.Y) * (p2.X - p1.X)
  out_p.X = p1.X - (p3.Y - p2.Y) * t& / s&
   out_p.Y = p1.Y + (p3.X - p2.X) * t& / s&
   If out_point_no% > 0 Then
    Call set_point_coordinate(out_point_no%, out_p, is_change)
   End If
End If
End Sub

Function point_number(na As String) As Integer
'由点的名确定点号
Dim i%
For i = 1 To last_conditions.last_cond(1).point_no '***
If m_poi(i).data(0).data0.name = na Then
point_number = i
Exit Function
End If
Next i
'先输入点的名称,有系统设点号,
'last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'm_poi(last_conditions.last_cond(1).point_no).data(0).data0.name = na
'point_number = last_conditions.last_cond(1).point_no
point_number = 0
End Function


Sub set_Divide_Point(ByVal p1%, ByVal p2%, ratio As Single, p%, is_draw As Boolean) '12.30
'设比例用于作图
'zuo
Dim l%
'On Error GoTo divide_point_error
  Dim temp_record As record_type
 If ratio = 0 Then
  Exit Sub
 End If
 l% = line_number(p1%, p2%, pointapi0, pointapi0, _
                  depend_condition(point_, p1%), depend_condition(point_, p2%), _
                  condition, condition_color, 1, 0)
If p% = 0 Then
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1 '***
'MDIForm1.Toolbar1.Buttons(21).Image = 33
 p% = last_conditions.last_cond(1).point_no '***
  'Call set_point_color(p%, 0)
End If
    m_poi(p%).data(0).degree = 0
    Call set_point_visible(p%, 1, False)
t_coord = add_POINTAPI(m_poi(p1%).data(0).data0.coordinate, _
             divide_POINTAPI_by_number(minus_POINTAPI(m_poi(p2%).data(0).data0.coordinate, _
              m_poi(p1%).data(0).data0.coordinate), ratio))
't_coord.Y = (num2% * m_poi(p1%).data(0).data0.coordinate.Y + num1% * m_poi(p2%).data(0).data0.coordinate.Y) / (num1% + num2%)
If is_draw Then
Call set_point_coordinate(p%, t_coord, False)
Else
Call set_point_coordinate(p%, t_coord, True)
End If
 record_0.data0.condition_data.condition_no = 0
             Call add_point_to_line(p%, l%, 0, display, is_draw, 0)
 '             Call draw_point(Draw_form, poi(p%), 0, display)
divide_point_error:
End Sub

Sub ratio_point1(p1%, p2%, num1%, num2%, p%, is_change As Boolean)
t_coord.X = (m_poi(p1%).data(0).data0.coordinate.X * num2% + m_poi(p2%).data(0).data0.coordinate.X * num1%) / (num1% + num2%)
t_coord.Y = (m_poi(p1%).data(0).data0.coordinate.Y * num2% + m_poi(p2%).data(0).data0.coordinate.Y * num1%) / (num1% + num2%)
Call set_point_coordinate(p%, t_coord, is_change)
End Sub

Function read_three_point_circle(T_mouse_down_coord As POINTAPI, read_circle_no%, draw_ty As Integer) As Boolean
'draw_ty=2
 Dim same_circle_no%
 Dim i%
 Dim read_out_type As Integer
 If draw_step = 0 Or draw_step = 3 Then
  operat_is_acting = True
 '选圆
 temp_circles_for_draw(0) = 0
 temp_lines_for_draw(0) = 0
read_out_type = draw_new_point(T_mouse_down_coord, ele1, ele2, red, False, 0)
      If read_out_type = new_point_on_circle Then '鼠标落在圆上
         temp_circles_for_draw(0) = 1 '选中一圆
          temp_circles_for_draw(1) = ele1.no '圆的序号
          read_three_point_circle = True
           Call redraw_point(temp_point(draw_step).no, condition_color, pointapi0)
            draw_step = draw_step + 2
      ElseIf read_out_type = new_point_on_line And draw_ty = 1 Then
         temp_lines_for_draw(0) = 1
          temp_lines_for_draw(1) = ele1.no
           read_three_point_circle = True
             Call redraw_point(temp_point(draw_step).no, condition_color, pointapi0)
              draw_step = draw_step + 2
      ElseIf read_out_type = exist_point Then
        Call read_circles_and_lines_from_point(temp_point(draw_step).no) '将与确定点关连的圆收集在 temp_circles_for_draw(1)
      '******************************************************************************************************************
      If temp_circles_for_draw(0) = 1 And temp_lines_for_draw(0) = 1 Then
           temp_circle(read_circle_no%) = temp_circles_for_draw(1) '圆的序号
           temp_line(read_circle_no%) = temp_lines_for_draw(1)
            Call C_display_picture.set_m_circle_color(temp_circle(read_circle_no%), conclusion_color)
          If draw_ty = 3 Then
             temp_line(read_circle_no%) = 0
             read_three_point_circle = True
              Call redraw_point(temp_point(draw_step).no, condition_color, pointapi0)
             draw_step = draw_step + 2
          ElseIf draw_ty = 4 Then
             temp_line(read_circle_no%) = 0
             read_three_point_circle = True
             Call redraw_point(temp_point(draw_step).no, condition_color, pointapi0)
             draw_step = draw_step + 2
          ElseIf draw_ty = 5 Then
          read_three_point_circle = True
            Call redraw_point(temp_point(draw_step).no, condition_color, pointapi0)
             draw_step = draw_step + 2
          ElseIf draw_ty = 6 Then
              Call C_display_picture.set_m_line_color(temp_line(0), conclusion_color)
          End If
     ElseIf temp_circles_for_draw(0) = 1 Then
          temp_circle(read_circle_no%) = temp_circles_for_draw(1) '圆的序号
            Call C_display_picture.set_m_circle_color(temp_circle(read_circle_no%), conclusion_color)
             read_three_point_circle = True
              Call redraw_point(temp_point(draw_step).no, condition_color, pointapi0)
             draw_step = draw_step + 2
     ElseIf temp_lines_for_draw(0) = 1 Then
          If draw_ty = 3 Or draw_ty = 4 Then
             draw_step = draw_step - 1
            read_three_point_circle = False
             Exit Function
          ElseIf draw_ty = 5 Or draw_ty = 6 Then
              temp_line(read_circle_no%) = temp_lines_for_draw(1) '圆的序号
              'temp_lines_for_draw(0) = 0
              Call C_display_picture.set_m_line_color(temp_line(0), conclusion_color)
              Call redraw_point(temp_point(draw_step).no, condition_color, pointapi0)
               draw_step = draw_step + 2
               
          End If
     End If
  Else
           draw_step = draw_step - 1
            read_three_point_circle = False
             Exit Function
  End If
 ElseIf draw_step = 1 Or draw_step = 4 Then
         Call set_select_point
         If draw_ty > 0 Then
           Call set_select_line(temp_lines_for_draw(1), temp_lines_for_draw(2), temp_lines_for_draw(3), temp_lines_for_draw(4))
         Else
          Call set_select_line
         End If
         Call set_select_circle(temp_circles_for_draw(1), temp_circles_for_draw(2), temp_circles_for_draw(3), temp_circles_for_draw(4))
         Call set_forbid_point(temp_point(draw_step - 1).no)
         Call set_forbid_line
         Call set_forbid_circle
    read_out_type = draw_new_point(T_mouse_down_coord, ele1, ele2, red, False, 255)
    'read_out_type = read_inter_point(T_mouse_down_coord, ele1, ele2, temp_point(draw_step).no, _
                           False, False)
       If read_out_type = exist_point Then
         If find_same_point_from_two_points(m_poi(temp_point(draw_step).no).data(0).in_circle, _
                         temp_circles_for_draw, temp_circles_for_draw, temp_point(draw_step).no) = 1 Then 'And _
                        '  (draw_ty = 0 Or find_same_point_from_two_points(m_poi(temp_point(draw_step).no).data(0).in_line, _
                        '    temp_lines_for_draw, temp_lines_for_draw, 0) = 0) Then
                         temp_circle(read_circle_no%) = temp_circles_for_draw(1)
         Call C_display_picture.set_m_circle_color(temp_circle(read_circle_no%), conclusion_color)
                  read_three_point_circle = True
                   Call redraw_point(temp_point(draw_step - 1).no, condition_color, pointapi0)
                   Call redraw_point(temp_point(draw_step).no, condition_color, pointapi0)
                    draw_step = draw_step + 1
         ElseIf temp_lines_for_draw(0) = 1 And temp_circles_for_draw(0) = 0 Then
                  temp_line(read_circle_no%) = temp_lines_for_draw(1)
                  Call C_display_picture.set_m_line_color(temp_line(0), conclusion_color)
                  Call redraw_point(temp_point(draw_step - 1).no, condition_color, pointapi0)
                   Call redraw_point(temp_point(draw_step).no, condition_color, pointapi0)
                 draw_step = draw_step + 1 ' 完成选圆
                  read_three_point_circle = True

         End If
      'ElseIf chose_one_circle_center() Then
      Else
              draw_step = draw_step - 1
      End If
 ElseIf draw_step = 2 Or draw_step = 5 Then
         Call set_select_point
         If draw_ty > 0 Then
           Call set_select_line(temp_lines_for_draw(1), temp_lines_for_draw(2), temp_lines_for_draw(3), temp_lines_for_draw(4))
         Else
          Call set_select_line
         End If
         Call set_select_circle(temp_circles_for_draw(1), temp_circles_for_draw(2), temp_circles_for_draw(3), temp_circles_for_draw(4))
         Call set_forbid_point(temp_point(draw_step - 2).no, temp_point(draw_step - 1).no)
         Call set_forbid_line
         Call set_forbid_circle
       If draw_new_point(mouse_down_coord, ele1, ele2, red, False, 1) > 0 Then '读出已知点
         If find_same_point_from_two_points(m_poi(temp_point(draw_step).no).data(0).in_circle, _
                                       temp_circles_for_draw, temp_circles_for_draw, temp_point(draw_step).no) >= 1 Then 'And
                          '(draw_ty = 0 Or find_same_point_from_two_points(m_poi(temp_point(draw_step).no).data(0).in_line, _
                            temp_lines_for_draw, temp_lines_for_draw, 0) = 0) Then
                         temp_circle(read_circle_no%) = temp_circles_for_draw(1)
                         read_three_point_circle = True
                    Call C_display_picture.set_m_circle_color(temp_circle(read_circle_no%), conclusion_color)
                    Call redraw_point(temp_point(draw_step - 2).no, condition_color, pointapi0)
                   Call redraw_point(temp_point(draw_step - 1).no, condition_color, pointapi0)
                   Call redraw_point(temp_point(draw_step).no, condition_color, pointapi0)
         ElseIf temp_lines_for_draw(0) = 1 And temp_circles_for_draw(0) = 0 Then
              temp_line(read_circle_no%) = temp_lines_for_draw(1)
              Call C_display_picture.set_m_line_color(temp_line(0), conclusion_color)
                   Call redraw_point(temp_point(draw_step - 2).no, condition_color, pointapi0)
                   Call redraw_point(temp_point(draw_step - 1).no, condition_color, pointapi0)
                   Call redraw_point(temp_point(draw_step).no, condition_color, pointapi0)
             read_three_point_circle = True
         End If
      Else
        draw_step = draw_step - 1
      End If
 End If
End Function
Function read_circle(last_point%, in_coord As POINTAPI, point_no%, _
                        out_coord As POINTAPI, control_ty As Byte, Optional is_set_data As Boolean = False) As Integer
'******************************************************
'从st_point%起读出（x!,y!）附近圆上的点（out_x%,out_x%）
'******************************************************
Dim s!
Dim i%
'On Error GoTo read_circle_error
read_circle = read_tangent_circle(in_coord, out_coord, point_no%, is_set_data)
If read_circle > 0 Then
 
 is_set_data = False
 Exit Function
Else
 read_circle = 0
End If
For i% = last_point% To 1 Step -1 'last_conditions.last_cond(1).circle_no  '***
 If find_point_from_points(point_no%, m_Circ(i%).data(0).data0.in_point) And point_no% > 0 Then '是圆上已有点
    GoTo mark1
 End If '   Exit Function 'GoTo mark1
 If m_Circ(i%).data(0).data0.visible > 0 Then
  If read_circle0(m_Circ(i%).data(0).data0, in_coord, out_coord) Then
    If get_control_data(circle_, i%, control_ty) Then
      read_circle = i%
       Exit Function
    End If
  End If
 End If
mark1:
Next i%
Exit Function
read_circle_error:
 read_circle = 0
End Function
Function read_inter_point(in_coord As POINTAPI, ele1 As condition_type, ele2 As condition_type, _
                            p%, ByVal is_set_data As Boolean, input_type As Integer, Optional need_control As Byte = 0) As Integer
                           'in_coord 输入点的坐标
                           'ele1,ele1 是选择的线圆号,是交点
                           'ele_ty,选择在线或圆上的点出
                           'ele_ty=point_ ele_no=0 读出所有的点,
                           'ele_ty=point_ ele_no=-1 只读新点,
                           'need_control =0 无控制
                           'need_control =1 必需落在已知点
                           'need_control =2 必需落在已知线
                           'need_control =4 必需落在已知圆
                           'need_control=5'必须落在切线上
                           'need_control =255 有控制
                           '*************************************************
                           'exist_point 已有点
                           '  line_ 选中第一线
                           ' interset_point_line_line   '两线相交于新点,交点靠近输入位置
                           'new_point_on_line_circle
                           'set_new_wenti_cond
                           'new_free_point
                           ' new_point_on_line
                           ' read_inter_point = 0'不成功的操作
                           'need_control=1 读出已有点true否则false
Dim t_p%, l%, t_i%, t_l%
Dim D1&, D2&
Dim t_last_point_no As Integer
Dim l_p(3) As Integer
Dim tl(1) As line_data_type
Dim vf As POINTAPI
Dim t_last_ele_no%
Dim aid_coord As POINTAPI
Dim in_point_no%
' 单点 1
t_p% = p%
ele1.ty = 0
ele2.ty = 0
ele1.no = 0
ele2.no = 0
ele1.old_no = 0
ele2.old_no = 0
' 线圆 5 , 7 已有一交点,且选中,8,9 ,13 负点
                                       ' 圆圆 6 ,10,11,12   ,14

t_last_point_no = last_conditions.last_cond(1).point_no '最后一点的序号
    ele1.no = read_point(in_coord, 0) '读 点
      p% = ele1.no
      in_point_no% = ele1.no '记录原始输入点的序号，
      ele1.ty = point_
      ele1.old_no = ele1.no
'***********************************************************************************
     If ele1.no > 0 Then  '    读中-点
      If get_control_data(point_, p%, need_control) = False Then
          read_inter_point = 0
          Exit Function
      Else 'If need_control = 1 Then
          read_inter_point = exist_point
          Exit Function
      End If
     Else
      ele1.ty = 0
'******************************************************************************
     End If 'ele1.no > 0 Then
 '****************************************************************************
'***************************************************************************************
'**************************************************************************************
 ' Else '未读中已有点
'************************************************************************************
t_last_ele_no% = C_display_picture.m_line.Count
read_inter_point_back0:
     ele1.no = read_line(t_last_ele_no%, in_coord, p%, t_coord1, is_set_data, input_type, need_control)        '读线
        If ele1.no <> 0 Then       ' 选中第一线
'********************************************************************************************
              ele1.ty = line_
              ele1.old_no = ele1.no
              read_inter_point = line_
'*********************************************************************************************
         l_p(0) = m_lin(ele1.no).data(0).data0.poi(0)
         l_p(1) = m_lin(ele1.no).data(0).data0.poi(1)
          tl(0) = m_lin(ele1.no).data(0)
          t_i% = ele1.no
'**********************************************************************************************
read_iter_point_back1:
'***************************************************************************************
'*************************************************************************************
      '选第二线
 '***********************************************************************************
          If ele1.no > 0 Then
             l% = ele1.no - 1
          Else
             l% = ele1.no + 1
          End If
 '**************************************************************************************
read_line_back:
            ele2.no = C_display_picture.read_line(l%, in_coord.X, in_coord.Y, p%, t_coord2.X, t_coord2.Y, is_set_data, need_control) '    再选线
  '**********************************************************************************
             If ele2.no <> 0 Then    '选中两线
  '*********************************************************************************
                If ele2.no = ele1.no Then
                 l% = l% + 1
                 GoTo read_line_back
                End If
  '***************************************************************************************
                ele2.ty = line_
                ele2.old_no = ele2.no
 '*******************************************************************************************
              tl(1) = m_lin(ele2.no).data(0)
               If p% = 0 Then
                    read_inter_point = interset_point_line_line   '两线相交于新点,交点靠近输入位置
                  If is_set_data Then
                   read_inter_point = inter_point_line_line(0, 0, 0, 0, ele1.no, ele2.no, _
                                   p%, in_coord, is_set_data, True, False)                        '两线交点
                    'm_poi(p%).data(0).degree = 0
                                     If is_same_POINTAPI(m_poi(p%).data(0).data0.coordinate, in_coord) Then
                     in_coord = m_poi(p%).data(0).data0.coordinate '
                       Call distance_point_to_line(in_coord, m_poi(l_p(0)).data(0).data0.coordinate, paral_, _
                             m_poi(l_p(0)).data(0).data0.coordinate, m_poi(l_p(1)).data(0).data0.coordinate, D1&, vf, 1)
                       Call distance_point_to_line(in_coord, m_poi(l_p(2)).data(0).data0.coordinate, paral_, _
                            m_poi(l_p(2)).data(0).data0.coordinate, m_poi(l_p(3)).data(0).data0.coordinate, D2&, vf, 1)
                    If Abs(D1&) > Abs(D2&) Then
                      ele2.ty = 0
                      t_coord1 = t_coord2
                       in_coord = t_coord1
                    End If
             '    Call set_point_parent(line_, ele1.no, m_poi(p%).data(0))
             '    Call set_point_parent(line_, ele2.no, m_poi(p%).data(0))
                 m_poi(p%).data(0).parent.inter_type = new_inter_point_type(m_poi(p%).data(0).parent.inter_type, interset_point_line_line)
                   End If 'if is_ste_data
               ElseIf p% > 0 Then '
                 l_p(0) = is_point_in_points(p%, m_lin(ele1.no).data(0).data0.in_point)
                 l_p(1) = is_point_in_points(p%, m_lin(ele2.no).data(0).data0.in_point)
                 If l_p(0) = 0 And l_p(1) > 0 And is_set_data Then '落在第二条线上
                      Call add_point_to_line(p%, ele1.no, 0, True, True, 0)
                      read_inter_point = new_point_on_line
                 ElseIf l_p(0) > 0 And l_p(1) = 0 And is_set_data Then '
                      ele1.no = ele2.no
                      Call add_point_to_line(p%, ele1.no, 0, True, True, 0)
                      read_inter_point = new_point_on_line
                 Else
                      read_inter_point = exist_point
                 End If
               End If 'if p%=0
          End If
      '*******************************************************************
      Else '未选中第二条线
      '*******************************************************************
     t_last_ele_no% = last_conditions.last_cond(1).circle_no
     ele2.no = read_circle(t_last_ele_no%, in_coord, p%, t_coord2, need_control) '  选圆
     '********************************************************************
        If ele2.no > 0 Then
           ele2.ty = circle_
           ele2.old_no = ele2.no
point_on_line_circle:
           If p% = 0 Then
                read_inter_point = new_point_on_line_circle
            If is_set_data Then 't_p%
              read_inter_point = _
                inter_point_line_circle(ele1.no, ele2.no, in_coord, p%, True, is_set_data)  ', t_p2%)
                 m_poi(p%).data(0).degree = 0
                  If read_inter_point >= new_point_on_line_circle Then 'And _
                    read_inter_point <= new_point_on_Tline_Tcircle21 Then
                      Call set_point_visible(p%, 1, False)
                        in_coord = m_poi(p%).data(0).data0.coordinate
                   End If
                'Call set_parent(ele1.ty, ele1.no, point_, p%, 0)
                'Call set_parent(ele2.ty, ele2.no, point_, p%, 0)
               ' m_poi(p%).data(0).parent.inter_type = new_inter_point_type(m_poi(p%).data(0).parent.inter_type, read_inter_point)
            End If 'is_ste_data
           ElseIf p% > 0 Then
                 l_p(0) = is_point_in_points(p%, m_lin(ele1.no).data(0).data0.in_point)
                 l_p(1) = is_point_in_points(p%, m_Circ(ele2.no).data(0).data0.in_point)
                 If l_p(0) = 0 And l_p(1) > 0 And is_set_data Then '落在第二条线上
                      Call add_point_to_line(p%, ele1.no, 0, True, True, 0)
                      read_inter_point = new_point_on_line
                 ElseIf l_p(0) > 0 And l_p(1) = 0 And is_set_data Then '
                      ele1.no = ele2.no
                      ele1.ty = circle_
                       Call add_point_to_m_circle(p%, ele1.no, record0, True)
                      read_inter_point = new_point_on_circle
                 Else
                      read_inter_point = exist_point
                 End If
          End If 'if p5=0
          'End If
      '********************************************************************************
        Else 'if ele2.no=0              '未选中一圆,线上的点
      '********************************************************************************
             'If is_special_geometry_ele(in_coord, ele1, ele2, need_control) = False Then
             '    read_inter_point = 0
             '    Exit Function
             'Else
        '***************************************************************************
              'If tl(0).data0.visible = 1 And tl(0).data0.in_point(0) = 3 Then
               'in_coord = t_coord1
               ' End If
                'If is_set_data Then
                 'MDIForm1.Toolbar1.Buttons(21).Image = 33 '新增点
                 'If ele1.no > last_conditions.last_cond(1).line_no Then
                 '   Call set_line_from_aid_line(ele1.old_no, p%, t_coord1, ele1.no)
                 'End If
                 If p% = 0 Then
                    read_inter_point = new_point_on_line
                  If is_set_data Then 'And aid_line_for_input = 0 And aid_circle_for_input = 0 Then
                    p% = m_point_number(t_coord1, condition, 1, condition_color, "", _
                     depend_condition(ele1.ty, ele1.no), ele2, new_point_on_line, True)
                      '‘m_poi(p%).data(0).degree = 1
                                          record_0.data0.condition_data.condition_no = 0
                       Call add_point_to_line(p%, ele1.no, 0, display, True, 0)
                        read_inter_point = new_point_on_line
                  End If
                 ElseIf p% > 0 Then
                   If m_poi(p%).data(0).parent.inter_type > 0 Then
                      read_inter_point = tangent_point_
                   Else
                      read_inter_point = exist_point
                   End If
                 End If 'if p%=0
        End If
        End If
 '***************************************************************************************
  Else '    If is_set_data Then '未选中线'选圆
 '****************************************************************************************
t_last_ele_no% = last_conditions.last_cond(1).circle_no
  ele1.no = read_circle(t_last_ele_no%, in_coord, p%, t_coord1, need_control, is_set_data) '选第一圆
    If ele1.no > 0 Then 'And is_set_data Then
       ele1.ty = circle_
     ele2.no = read_circle(ele1.no - 1, in_coord, p%, t_coord2, need_control) '选第二圆
      If ele2.no > 0 Then '
         ele2.ty = circle_
         ele2.old_no = ele2.no
'************************************************
'圆圆上的点
'**************************************************
point_on_circle_circle:
   read_inter_point = inter_point_circle_circle(ele1.no, ele2.no, in_coord)
 '************************************************************8
                If p% = 0 And is_set_data Then
                  p% = m_point_number(in_coord, condition, 1, condition_color, "", ele1, ele2, read_inter_point, True)
                   m_poi(p%).data(0).degree = 0
                   Call add_point_to_m_circle(p%, ele1.no, record0, True) '= 0 Then '已有点
                   Call add_point_to_m_circle(p%, ele2.no, record0, True) '= 0 Then '已有点
                End If
                m_poi(p%).data(0).parent.inter_type = new_inter_point_type(m_poi(p%).data(0).parent.inter_type, read_inter_point)
      Else 'ele2=0圆上的点
           in_coord = t_coord1
                       read_inter_point = new_point_on_circle
                                           read_inter_point = new_point_on_circle     '在单圆上
                       If p% = 0 And is_set_data Then
                         p% = m_point_number(in_coord, condition, 1, condition_color, "", ele1, ele2, new_point_on_circle, True)
                          m_poi(p%).data(0).degree = 1
                            Call add_point_to_m_circle(p%, ele1.no, record0, True) '= 0 Then '已有点
                       End If
                'Exit Function
      End If
      'End If
      'End If
  Else
  '**************************************************
    If p% > 0 Then
                     read_inter_point = exist_point
    ElseIf t_p% = 0 And is_set_data Then
       If is_special_geometry_ele(in_coord, ele1, ele2, need_control) Then
          p% = m_point_number(in_coord, condition, 1, condition_color, "", ele1, ele2, new_free_point, True)
                    read_inter_point = new_free_point
       Else
                       read_inter_point = 0
       End If
              ele1.ty = 0
              ele1.no = 0
              ele1.old_no = 0
              ele2.ty = 0
              ele2.no = 0
              ele2.old_no = 0
      End If
     ' Exit Function
   End If
  End If
End Function

Sub put_point_on_circle(in_x%, in_y%, c%, out_x%, out_y%)
Dim s!
s! = sqr((in_x% - m_Circ(c%).data(0).data0.c_coord.X) ^ 2 + _
         (in_y% - m_Circ(c%).data(0).data0.c_coord.Y) ^ 2)
   If s! = 0 Then
    Exit Sub
   End If
 out_x% = m_Circ(c%).data(0).data0.c_coord.X + _
    (in_x% - m_Circ(c%).data(0).data0.c_coord.X) * m_Circ(c%).data(0).data0.radii / s!
 out_y% = m_Circ(c%).data(0).data0.c_coord.Y + _
    (in_y% - m_Circ(c%).data(0).data0.c_coord.Y) * m_Circ(c%).data(0).data0.radii / s!

End Sub

Public Sub redraw_three_point_circle(c As circle_data0_type, is_change As Boolean)
Dim X(1) As Single
Dim Y(1) As Single
Dim z(1) As Single
Dim temp_num0!, temp_num1!, temp_num2!
X(0) = CSng(m_poi(c.in_point(1)).data(0).data0.coordinate.X)
X(1) = CSng(m_poi(c.in_point(1)).data(0).data0.coordinate.Y)
Y(0) = CSng(m_poi(c.in_point(2)).data(0).data0.coordinate.X)
Y(1) = CSng(m_poi(c.in_point(2)).data(0).data0.coordinate.Y)
z(0) = CSng(m_poi(c.in_point(3)).data(0).data0.coordinate.X)
z(1) = CSng(m_poi(c.in_point(3)).data(0).data0.coordinate.Y)
        temp_num0! = X(0) * Y(1) + Y(0) * z(1) + z(0) * X(1) - X(0) * z(1) - Y(0) * X(1) - z(0) * Y(1)
         If temp_num0! = 0 Then
         Exit Sub
         End If

           temp_num1! = (X(1) - Y(1)) * (Y(0) ^ 2 - z(0) ^ 2) - (Y(1) - z(1)) * (X(0) ^ 2 - Y(0) ^ 2)
            temp_num2! = (X(1) - Y(1)) * (Y(1) - z(1)) * (z(1) - X(1))
         t_coord.X = CInt(-(temp_num1! + temp_num2!) / temp_num0! / 2)
           temp_num1! = (X(0) - Y(0)) * (Y(1) ^ 2 - z(1) ^ 2) - (Y(0) - z(0)) * (X(1) ^ 2 - Y(1) ^ 2)
            temp_num2! = (X(0) - Y(0)) * (Y(0) - z(0)) * (z(0) - X(0))
         t_coord.Y = CInt((temp_num1! + temp_num2!) / (temp_num0! * 2))
           c.radii = sqr((X(0) - m_poi(c.center).data(0).data0.coordinate.X) ^ 2 + (X(1) - m_poi(c.center).data(0).data0.coordinate.Y) ^ 2)
         Call set_point_coordinate(c.center, t_coord, is_change)
End Sub



'在重画过程中组织结论线段
'
'
Public Sub simple_con_line(l As line_data0_type)
Dim i%, j%, k%, m%, n%
Dim tl As line_data0_type
If (m_poi(l.poi(0)).data(0).data0.coordinate.X = 10000 And _
      m_poi(l.poi(0)).data(0).data0.coordinate.Y = 10000) Or _
       (m_poi(l.poi(1)).data(0).data0.coordinate.X = 10000 And _
         m_poi(l.poi(1)).data(0).data0.coordinate.Y = 10000) Then
Exit Sub
End If
 For i% = 0 To l.in_point(0)
  tl.in_point(i%) = Abs(l.in_point(i%))
 Next i%
  'Call simple_line(tl)
 For i% = 1 To l.in_point(0)
  If l.in_point(i%) < 0 Then
   For j% = 1 To l.in_point(0)
    If tl.in_point(j%) = -l.in_point(i%) Then
      For k% = i% + 1 To l.in_point(0)
      If l.in_point(k%) > 0 Then
        For m% = 1 To l.in_point(0)
         If l.in_point(k%) = tl.in_point(m%) Then
          If m% < j% Then
           For n% = m% To j% - 1
            tl.in_point(n%) = -tl.in_point(n%)
           Next n%
          ElseIf j% < m% Then
           For n% = j% To m% - 1
            tl.in_point(n%) = -tl.in_point(n%)
           Next n%
          End If
         End If
        Next m%
       k% = i%
      GoTo simple_con_line_next
      End If
     Next k%
    End If
   Next j%
  End If
simple_con_line_next:
 Next i%
   If tl.in_point(tl.in_point(0) - 1) > 0 Then
     tl.in_point(tl.in_point(0)) = -tl.poi(1)
   End If
   l = tl
End Sub

Public Function calculate_line_line_intersect_point(p1 As POINTAPI, _
          p2 As POINTAPI, p3 As POINTAPI, p4 As POINTAPI, _
            inter_point As POINTAPI, is_change As Boolean) As Boolean        '判断两直线（四个端点）是否相交，并计算两直线的交点坐标
Dim d&, A&
Dim r!
d& = (p2.X - p1.X) * (p4.Y - p3.Y) - _
       (p2.Y - p1.Y) * (p4.X - p3.X) '
If d& = 0 Then
     calculate_line_line_intersect_point = False
Else
A& = (p3.X - p1.X) * (p4.Y - p3.Y) - _
        (p3.Y - p1.Y) * (p4.X - p3.X)
r! = A& / d&
   inter_point.X = p1.X + (p2.X - p1.X) * r!
   inter_point.Y = p1.Y + (p2.Y - p1.Y) * r!
     calculate_line_line_intersect_point = True
End If
End Function
Public Function inter_point_line_line2(coord1 As POINTAPI, t1 As Integer, _
            coord2 As POINTAPI, coord3 As POINTAPI, coord4 As POINTAPI, t2 As Integer, _
             coord5 As POINTAPI, coord6 As POINTAPI, out_coord As POINTAPI, out_p%, _
               is_change As Boolean, is_set_data As Boolean) As Byte
Dim A&, b&, c&, d&, p&, q&, s&, t!
If (coord3.X - coord2.X = 0 And coord3.Y - coord2.Y = 0) Or _
       (coord6.X - coord5.X = 0 And coord6.Y - coord5.Y = 0) Then
     Exit Function
End If
 p& = coord4.X - coord1.X
 q& = coord4.Y - coord1.Y
If t1 = paral_ Then
 A& = coord3.X - coord2.X
 c& = coord3.Y - coord2.Y
ElseIf t1 = verti_ Then
 A& = coord3.Y - coord2.Y
 c& = coord2.X - coord3.X
End If
If t2 = paral_ Then
  b& = coord5.X - coord6.X
  d& = coord5.Y - coord6.Y
ElseIf t2 = verti_ Then
  b& = coord5.Y - coord6.Y
  d& = coord6.X - coord5.X
End If
s& = (A& * d& - c& * b&)
If s& = 0 Then
t! = 20000 / (Abs(coord3.X - coord2.X) + Abs(coord3.Y - coord2.Y))
 inter_point_line_line2 = False
Else
t! = (p& * d& - q& * b&) / s& ' (A& * d& - c& * b&)
 inter_point_line_line2 = True
End If
 out_coord.X = coord1.X + A& * t!
 out_coord.Y = coord1.Y + c& * t!
 If out_p% > 0 Then
  Call set_point_coordinate(out_p%, out_coord, is_change)
 ElseIf is_set_data Then
  out_p% = m_point_number(out_coord, condition, 1, condition_color, "", condition_type0, condition_type0, 0, is_set_data)
  inter_point_line_line2 = 1
 End If
End Function

Public Function inter_point_line_line3(ByVal p%, ByVal ty As Integer, _
                       ByVal l1%, ByVal p1%, ByVal ty1 As Integer, ByVal l2%, _
                        out_coord As POINTAPI, out_p%, is_change As Boolean, _
                         c_data As condition_data_type, is_set_data As Boolean) As Byte               '新交点=1,2 推出结论
'过p%点平行(垂直)l1%的直线交过p1%点平行(垂直)l2%的直线于(out_coord)
Dim i%, j%, tp%, tl%, out_l1%, out_l2%
Dim tl_(1) As Integer
'p%在l1%上,过p%点平行(垂直)l1%的直线=l1%
Call simple_line_from_point_line(p%, ty, l1%)
Call simple_line_from_point_line(p1%, ty1, l2%)
'***************************************************************************
If p% = 0 And p1% = 0 Then 'p%=0,p1%=0,l两线交点
    inter_point_line_line3 = inter_point_line_line(0, 0, 0, 0, l1%, l2%, out_p%, _
        out_coord, is_set_data, True, is_change)
        out_l1% = l1%
        out_l2% = l2%
              If is_set_data And inter_point_line_line3 = 1 Then
                    inter_point_line_line3 = set_inter_point_line_line_data(p%, ty, l1%, p1%, _
                     ty1, l1%, out_l1%, out_l2%, 0, c_data)
              End If
        Exit Function
ElseIf p1% = 0 Then '过p%点平行p1%(垂直)的直线与l2%相交
    If ty = paral_ Then '平行
      If is_dparal(l1%, l2%, 0, -1000, 0, 0, 0, 0) Then
          Exit Function '
      Else
              inter_point_line_line3 = inter_point_line_line2(m_poi(p%).data(0).data0.coordinate, _
                ty, m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                    m_poi(m_lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate, _
                     m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                      paral_, m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                       m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate, out_coord, _
                         out_p%, is_change, is_set_data)
              out_l1% = line_number0(out_p%, p%, 0, 0)
              out_l2% = l2%
              If is_set_data And inter_point_line_line3 = 1 Then
                    inter_point_line_line3 = set_inter_point_line_line_data(p%, ty, l1%, p1%, _
                     ty1, l1%, out_l1%, out_l2%, out_p%, c_data)
              End If
      End If
     ElseIf ty = verti_ Then 'ty=0
      If is_dverti(l1%, l2%, 0, -1000, 0, 0, 0, 0) Then
            Exit Function
      Else
           inter_point_line_line3 = inter_point_line_line2(m_poi(p%).data(0).data0.coordinate, _
                ty, m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                    m_poi(m_lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate, _
                     m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                      paral_, m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                       m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate, out_coord, _
                        out_p%, is_change, is_set_data)
              out_l1% = line_number0(out_p%, p%, 0, 0)
              out_l2% = l2%
              If is_set_data And inter_point_line_line3 = 1 Then
                    inter_point_line_line3 = set_inter_point_line_line_data(p%, ty, l1%, p1%, _
                     ty1, l1%, out_l1%, out_l2%, out_p%, c_data)
              End If
      End If
    End If 'ty
ElseIf p% = 0 Then
    inter_point_line_line3 = inter_point_line_line3(p1%, ty1, l2%, p%, ty, l1%, _
                               out_coord, out_p%, is_change, c_data, is_set_data)
Else 'p%>0 and p1%>0
       inter_point_line_line3 = inter_point_line_line2(m_poi(p%).data(0).data0.coordinate, ty, _
            m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate, _
              m_poi(m_lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate, _
                 m_poi(p1%).data(0).data0.coordinate, ty1, _
            m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate, _
              m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate, _
                               out_coord, out_p%, is_change, is_set_data)
               out_l1% = line_number0(out_p%, p%, 0, 0)
               out_l2% = line_number0(out_p%, p1%, 0, 0)
              If is_set_data And inter_point_line_line3 = 1 Then
                    inter_point_line_line3 = set_inter_point_line_line_data(p%, ty, l1%, p1%, _
                     ty1, l1%, out_l1%, out_l2%, out_p%, c_data)
              End If
End If
End Function
Public Function con_line_number(p1%, p2%) As Integer
Dim i%, j%, n%
Dim tp(1) As Integer
For i% = 1 To last_con_line1
 n% = 0
 For j% = 1 To con_line1(i%).data(0).data0.in_point(0)
  If con_line1(i%).data(0).data0.in_point(j%) = p1% Or _
   con_line1(i%).data(0).data0.in_point(j%) = p2% Then
    n% = n% + 1
     If n% = 2 Then
      con_line_number = i%
       Exit Function
      End If
   End If
  Next j%
   Next i%
last_con_line1 = last_con_line1 + 1
 'ReDim Preserve con_line1(last_con_line1).data(0) As line_data_type
 'ReDim Preserve con_line1(last_con_line1).data(1) As line_data_type
If compare_two_point(m_poi(p1%).data(0).data0.coordinate, _
           m_poi(p2%).data(0).data0.coordinate, 0, 0, 6) = 1 Then
 tp(0) = p1%
  tp(1) = p2%
Else
 tp(0) = p2%
  tp(1) = p1%
End If
con_line1(last_con_line1).data(0).data0.in_point(0) = 2
 con_line1(last_con_line1).data(0).data0.in_point(1) = tp(0)
  con_line1(last_con_line1).data(0).data0.in_point(2) = tp(1)
   con_line1(last_con_line1).data(0).data0.poi(0) = tp(0)
    con_line1(last_con_line1).data(0).data0.poi(1) = tp(1)
 con_line_number = last_con_line1
    
    

End Function

Public Sub Drawline(ob As Object, color As Long, line_ty%, p1 As POINTAPI, p2 As POINTAPI, dir, visible As Byte)
If (Abs(p1.X) = 10000 And Abs(p1.Y) = 10000) Or _
    (Abs(p2.X) = 10000 And Abs(p2.Y) = 10000) Or _
      visible = 0 Then
   Exit Sub
End If
If color = 0 Then '设置颜色
 color = QBColor(condition_color)
End If
If line_ty = 0 Then '类型0
ob.Line (p1.X, p1.Y)-(p2.X, p2.Y), color
If regist_data.run_type = 1 Then '向量
If dir = 1 Then '画箭头
 Call draw_arrow(ob, p1.X, p1.Y, p2.X, p2.Y, color)
ElseIf dir = -1 Then
 Call draw_arrow(ob, p2.X, p2.Y, p1.X, p1.Y, color)
End If
End If
ElseIf line_ty = 1 Then '类型1画虚线
ob.DrawStyle = 2
ob.Line (p1.X, p1.Y)-(p2.X, p2.Y), color
If regist_data.run_type = 1 Then
If dir = 1 Then
 Call draw_arrow(ob, p1.X, p1.Y, p2.X, p2.Y, color)
ElseIf dir = -1 Then
 Call draw_arrow(ob, p2.X, p2.Y, p1.X, p1.Y, color)
End If
End If
ob.DrawStyle = 0 '恢复作图板状态
End If
End Sub
Private Sub delete_select_control_data(ele_ty, ele_no%)
If ele_ty = point_ Then
   Call delete_point_from_points(ele_no, control_data.select_point)
ElseIf ele_ty = line_ Then
   Call delete_point_from_points(ele_no, control_data.select_line)
ElseIf ele_ty = circle_ Then
   Call delete_point_from_points(ele_no, control_data.select_circle)
End If
End Sub
Private Sub delete_point_from_points(in_p%, points() As Integer)
Dim i%, j%
For i% = 1 To points(0)
 If points(i%) = in_p% Then
    points(0) = points(0) - 1
    For j% = i% To points(0)
       points(i%) = points(i% + 1)
    Next j%
 End If
Next i%
End Sub
Public Function get_control_data(ele_ty As Byte, ele_no As Integer, control_ty As Byte) As Boolean
Dim i%
If control_ty = 255 Then
If ele_ty = point_ Then
   For i% = 1 To control_data.forbid_point(0)
      If control_data.forbid_point(i%) = ele_no Then '是禁用元素
          get_control_data = False '
           Exit Function
      End If
   Next i%
          get_control_data = True
   If control_data.select_point(0) = 0 Then '无规定元
          get_control_data = True
   Else
   For i% = 1 To control_data.select_point(0)
      If control_data.select_point(i%) = ele_no Then '选中规定元
        Call delete_select_control_data(point_, ele_no) '消去规定
          get_control_data = True '
           Exit Function
      End If
   Next i%
   End If
ElseIf ele_ty = line_ Then
   For i% = 1 To control_data.forbid_line(0)
      If control_data.forbid_line(i%) = ele_no Then
          get_control_data = False
           Exit Function
      End If
   Next i%
          get_control_data = True
    If control_data.select_line(0) = 0 Then
           get_control_data = True
    Else
   For i% = 1 To control_data.select_line(0)
      If control_data.select_line(i%) = ele_no Then
          Call delete_select_control_data(line_, ele_no)
          get_control_data = True
           Exit Function
      End If
   Next i%
    End If
ElseIf ele_ty = circle_ Then
   For i% = 1 To control_data.forbid_circle(0)
      If control_data.forbid_circle(i%) = ele_no Then
          get_control_data = False
           Exit Function
      End If
   Next i%
          get_control_data = True
   If control_data.select_circle(0) = 0 Then
          get_control_data = True
   Else
   For i% = 1 To control_data.select_circle(0)
      If control_data.select_circle(i%) = ele_no Then
         Call delete_select_control_data(circle_, ele_no)
          get_control_data = True
           Exit Function
      End If
   Next i%
   End If
Else
   get_control_data = True
End If
Else
   get_control_data = True
End If
End Function

Public Sub draw_circle1(ob As Object, c As circle_data0_type, ty)
If c.visible = 0 Then
Exit Sub
End If
If ty = condition Then
'If c.data(0).data0.color = 0 Then
'c.data(0).data0.color =condition_color
'End If
ob.Circle (c.c_coord.X, _
              c.c_coord.Y), c.radii, _
                  QBColor(condition_color)
Else
ob.Circle (c.c_coord.X, _
            c.c_coord.Y), c.radii, _
                  QBColor(conclusion_color)
End If
End Sub


Public Sub measur_again()
Dim i%
Dim n As Byte
Dim temp_s$
Dim sq_ratio_for_measur As Single
Dim vf As POINTAPI
For n = 0 To last_measur_string
 Call display_m_string(n, no_display)
Next n
If length_(0).poi(0) > 0 And length_(0).poi(1) > 0 Then
 Call draw_ruler(Ratio_for_measure.Ratio_for_measure, delete)
 Call length_of_line(length_(0))
  Ratio_for_measure.Ratio_for_measure = length_(0).len / length_(0).len0
  Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
  Call display_m_string(0, display)
ElseIf length_point_to_line(n).poi(0) > 0 And _
   length_point_to_line(n).poi(1) > 0 And length_point_to_line(n).poi(2) > 0 Then
 Call draw_ruler(Ratio_for_measure.Ratio_for_measure, delete)
Call distance_point_to_line( _
       m_poi(length_point_to_line(0).poi(0)).data(0).data0.coordinate, _
        m_poi(length_point_to_line(0).poi(1)).data(0).data0.coordinate, paral_, _
         m_poi(length_point_to_line(0).poi(1)).data(0).data0.coordinate, _
          m_poi(length_point_to_line(0).poi(2)).data(0).data0.coordinate, _
         length_point_to_line(0).len, vf, 1)
  Ratio_for_measure.Ratio_for_measure = Abs(length_point_to_line(n).len) / _
       length_point_to_line(n).len0
  Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
  Call display_m_string(0, display)
End If
For n = 1 To last_length
Call length_of_line(length_(n))
   Measur_string(n) = LoadResString_(1615, _
          "\\1\\" + m_poi(length_(n).poi(0)).data(0).data0.name + _
                  m_poi(length_(n).poi(1)).data(0).data0.name + _
          "\\2\\" + str_for_measure(length_(n).len / Ratio_for_measure.Ratio_for_measure))
Next n
'****
For n = 1 To last_length_point_to_line
Call distance_point_to_line( _
       m_poi(length_point_to_line(n).poi(0)).data(0).data0.coordinate, _
        m_poi(length_point_to_line(n).poi(1)).data(0).data0.coordinate, paral_, _
         m_poi(length_point_to_line(n).poi(1)).data(0).data0.coordinate, _
        m_poi(length_point_to_line(n).poi(2)).data(0).data0.coordinate, _
         length_point_to_line(n).len, vf)
    Measur_string(n + last_length) = LoadResString_(1785, _
           "\\1\\" + m_poi(length_point_to_line(n).poi(0)).data(0).data0.name + _
           "\\2\\" + m_poi(length_point_to_line(n).poi(1)).data(0).data0.name + _
                          m_poi(length_point_to_line(n).poi(2)).data(0).data0.name + _
           "\\3\\" + str_for_measure(Abs(length_point_to_line(n).len) / Ratio_for_measure.Ratio_for_measure))
Next n
'***
For n = 1 To last_angle_value_for_measur
 Call value_of_angle(angle_value_for_measur(n))

    Measur_string(n + last_length + last_length_point_to_line) = set_display_angle0( _
     m_poi(angle_value_for_measur(n).poi(0)).data(0).data0.name + _
      m_poi(angle_value_for_measur(n).poi(1)).data(0).data0.name + _
       m_poi(angle_value_for_measur(n).poi(2)).data(0).data0.name) + _
       "=" + angle_value_for_measur(n).value
Next n
sq_ratio_for_measur = Ratio_for_measure.Ratio_for_measure ^ 2
For n = 1 To last_Area_polygon
Area_polygon(n).Area = C_Area_polygon(Area_polygon(n).p)
 temp_s$ = ""
  For i% = 0 To Area_polygon(n).p.total_v - 1
   temp_s$ = temp_s$ + m_poi(Area_polygon(n).p.v(i%)).data(0).data0.name
  Next i%
     If Area_polygon(n).p.total_v = 3 Then
      temp_s$ = set_display_triangle0(temp_s$, 0, 0)
     Else
      temp_s$ = LoadResString_(1965, "\\1\\" + temp_s$)
     End If
     Measur_string(n + last_length + last_length_point_to_line + _
       last_angle_value_for_measur) = LoadResString_(1795, "\\1\\" + temp_s$ + _
          "\\2\\" + str_for_measure(Area_polygon(n).Area / sq_ratio_for_measur))
Next n
For n = 1 To last_measur_string
 Call display_m_string(n, display)
Next n

End Sub
Public Sub draw_ruler(ByVal m%, display_or_delete As Boolean)
Dim i%, max_size%, t%
If Ratio_for_measure.Ratio_for_measure = 0 Then '没有关于长度(或面积)的条件
  Exit Sub
End If
 Wenti_form.Picture3.CurrentX = 26 '设置当前位置
 Wenti_form.Picture3.CurrentY = 10
 Wenti_form.Picture3.FontSize = 10 '设置字体大小
'****************************************************************************************
If m% < 12 And m% > 0 Then '有ratio_for_measure和首次输入的长度条件确定标尺长
max_size% = CInt(12 / m%)
m% = m% * 10
t = 0.1
ElseIf m < 120 And m% > 0 Then
max_size% = CInt(120 / m%)
t% = 1
ElseIf m% > 0 Then
max_size% = CInt(1200 / m%)
t% = 10
m% = m% / 10
End If
'If display_or_delete = display Then
Call SetTextColor(Wenti_form.Picture3.hdc, QBColor(9))
'Else
'Call SetTextColor(Wenti_form.Picture3.hDC, QBColor(15))
'End If
Wenti_form.Picture3.Print LoadResString_(1970, "")
Wenti_form.Picture3.FontSize = 8.25
Wenti_form.Picture3.Line (26, 40)-(26 + m% * max_size, 40), QBColor(12)
For i% = 0 To max_size%
Wenti_form.Picture3.Line (26 + m% * i%, 38)-(26 + m% * i%, 41), QBColor(12)
Wenti_form.Picture3.CurrentY = 42
If t% = 1 Then
 Wenti_form.Picture3.CurrentX = 26 + m% * i% - 6
  Wenti_form.Picture3.Print i%
ElseIf t% = 10 Then
  If i% = 10 Or i% = 0 Then
   Wenti_form.Picture3.CurrentX = 26 + m% * i% - 6
    Wenti_form.Picture3.Print i% / 10
  Else
   Wenti_form.Picture3.CurrentX = 26 + m% * i% - 6
    Wenti_form.Picture3.Print ".";
     Wenti_form.Picture3.CurrentX = Wenti_form.Picture3.CurrentX - 4
    Wenti_form.Picture3.Print i%
  End If
Else
If i% = 0 Then
 Wenti_form.Picture3.CurrentX = 26 - 6
  Wenti_form.Picture3.Print i%
 Else
   Wenti_form.Picture3.CurrentX = 26 + m% * i% - 10
    Wenti_form.Picture3.Print i% * 10
 End If
End If
Next i%
Call SetTextColor(Wenti_form.Picture3.hdc, QBColor(0))
Call draw_coordianter(Wenti_form.Picture3)
End Sub
Public Sub draw_coordianter(ob As Object)
If is_set_function_data > 0 Then
  ob.Line (120, Wenti_form.Picture3.ScaleHeight - 200)- _
       (120, Wenti_form.Picture3.ScaleHeight - 580), QBColor(7)
  Call draw_arrow(ob, 120, Wenti_form.Picture3.ScaleHeight - 200, _
         120, Wenti_form.Picture3.ScaleHeight - 580, QBColor(7))
  ob.Line (120, Wenti_form.Picture3.ScaleHeight - 200)- _
       (480, Wenti_form.Picture3.ScaleHeight - 200), QBColor(7)
  Call draw_arrow(ob, 120, Wenti_form.Picture3.ScaleHeight - 200, _
        480, Wenti_form.Picture3.ScaleHeight - 200, QBColor(7))
End If
End Sub
Public Sub draw_tangent_line(n%, Optional display_or_delete As Byte = 0, Optional is_set_data As Boolean = False)
If display_or_delete = 1 And is_set_data = False Then '显示
If tangent_line(n%).data(0).visible = 0 Then
 Exit Sub
ElseIf tangent_line(n%).data(0).visible = 1 Then
   Call m_BPset(Draw_form, tangent_line(n%).data(0).coordinate(0), "", fill_color)
   Call m_BPset(Draw_form, tangent_line(n%).data(0).coordinate(1), "", fill_color)
   Draw_form.Line (tangent_line(n%).data(0).coordinate(0).X, tangent_line(n%).data(0).coordinate(0).Y)- _
    (tangent_line(n%).data(0).coordinate(1).X, tangent_line(n%).data(0).coordinate(1).Y), QBColor(fill_color)
    tangent_line(n%).data(0).visible = 2
ElseIf tangent_line(n%).data(0).visible = 2 Then
      Draw_form.Line (tangent_line(n%).data(0).old_coordinate(0).X, tangent_line(n%).data(0).old_coordinate(0).Y)- _
    (tangent_line(n%).data(0).old_coordinate(1).X, tangent_line(n%).data(0).old_coordinate(1).Y), QBColor(fill_color)
      tangent_line(n%).data(0).old_coordinate(0) = tangent_line(n%).data(0).new_coordinate(0)
      tangent_line(n%).data(0).old_coordinate(1) = tangent_line(n%).data(0).new_coordinate(1)
      Draw_form.Line (tangent_line(n%).data(0).old_coordinate(0).X, tangent_line(n%).data(0).old_coordinate(0).Y)- _
    (tangent_line(n%).data(0).old_coordinate(1).X, tangent_line(n%).data(0).old_coordinate(1).Y), QBColor(fill_color)
ElseIf tangent_line(n%).data(0).visible = 3 Then
     Call draw_tangent_line_by_ty(n%)
      tangent_line(n%).data(0).old_coordinate(0) = tangent_line(n%).data(0).new_coordinate(0)
      tangent_line(n%).data(0).old_coordinate(1) = tangent_line(n%).data(0).new_coordinate(1)
    Call draw_tangent_line_by_ty(n%)
ElseIf tangent_line(n%).data(0).visible = 4 Then
       tangent_line(n%).data(0).old_coordinate(0) = tangent_line(n%).data(0).new_coordinate(0)
       tangent_line(n%).data(0).old_coordinate(1) = tangent_line(n%).data(0).new_coordinate(1)
          Call draw_tangent_line_by_ty(n%)
       tangent_line(n%).data(0).visible = 5
ElseIf tangent_line(n%).data(0).visible = 5 Then
          Call draw_tangent_line_by_ty(n%)
          tangent_line(n%).data(0).old_coordinate(0) = tangent_line(n%).data(0).new_coordinate(0)
          tangent_line(n%).data(0).old_coordinate(1) = tangent_line(n%).data(0).new_coordinate(1)
          Call draw_tangent_line_by_ty(n%)
End If
 ElseIf display_or_delete = 0 Or is_set_data Then
 If tangent_line(n%).data(0).visible = 1 Then
       Call m_BPset(Draw_form, tangent_line(n%).data(0).coordinate(0), "", fill_color)
       Call m_BPset(Draw_form, tangent_line(n%).data(0).coordinate(1), "", fill_color)
      Draw_form.Line (tangent_line(n%).data(0).coordinate(0).X, tangent_line(n%).data(0).coordinate(0).Y)- _
      (tangent_line(n%).data(0).coordinate(1).X, tangent_line(n%).data(0).coordinate(1).Y), QBColor(fill_color)
 ElseIf tangent_line(n%).data(0).visible = 2 Then
     Call m_BPset(Draw_form, tangent_line(n%).data(0).coordinate(0), "", fill_color)
     Call m_BPset(Draw_form, tangent_line(n%).data(0).coordinate(1), "", fill_color)
      Draw_form.Line (tangent_line(n%).data(0).coordinate(0).X, tangent_line(n%).data(0).coordinate(0).Y)- _
    (tangent_line(n%).data(0).coordinate(1).X, tangent_line(n%).data(0).coordinate(1).Y), QBColor(fill_color)
      Draw_form.Line (tangent_line(n%).data(0).old_coordinate(0).X, tangent_line(n%).data(0).old_coordinate(0).Y)- _
    (tangent_line(n%).data(0).old_coordinate(1).X, tangent_line(n%).data(0).old_coordinate(1).Y), QBColor(fill_color)
     tangent_line(n%).data(0).old_coordinate(0).X = 10000
     tangent_line(n%).data(0).old_coordinate(0).Y = 10000
     tangent_line(n%).data(0).old_coordinate(1).X = 10000
     tangent_line(n%).data(0).old_coordinate(1).Y = 10000
ElseIf tangent_line(n%).data(0).visible = 3 Or tangent_line(n%).data(0).visible >= 4 Then
    Call draw_tangent_line_by_ty(n%)
End If
End If
End Sub

Public Function point_to_ratio(p1 As POINTAPI, p2 As POINTAPI, _
   p3 As POINTAPI, r1%)
Dim n1!

If Abs(p1.X - p3.X) > 5 Then
n1! = (p1.X - p2.X) / (p1.X - p3.X)
r1% = CInt(n1 * 1000)
point_to_ratio = True
ElseIf Abs(p1.Y - p3.Y) > 5 Then
n1! = (p1.Y - p2.Y) / (p1.Y - p3.Y)
r1% = CInt(n1 * 1000)
point_to_ratio = True
Else
r1% = 500
point_to_ratio = False
End If
End Function

Public Sub set_grey_line(ByVal l%)
Dim i%
For i% = 0 To 7
If grey_line(i%) = l% Then
Exit Sub
End If
Next i%
For i% = 0 To 7
If grey_line(i%) = 0 Then
grey_line(i%) = l%
Exit Sub
End If
Next i%

End Sub
Public Sub set_polygon(ByVal p1%, ByVal p2%, _
                            n%, polygon_no%, ByVal d As Boolean, is_e As Boolean, _
            num%)
 'n% =3,4,5,6
Dim i%, l%, j%, tp%
Dim tl(5) As Integer
Dim A!
Dim v$
Dim temp_record As total_record_type
If polygon_no% = 0 Then
For i% = 1 To last_conditions.last_cond(1).poly_no
    If poly(i%).total_v = n And poly(i%).is_e_polygon = is_e Then
       If poly(i%).v(0) = p1% And poly(i%).v(1) = p2% And poly(i%).direction = d Then
          polygon_no% = i%
           For j% = 2 To n% - 1
            If poly(i%).v(j%) = 0 Then
             GoTo set_polygon_mark0 '数据不完整
            End If
           Next j%
            Exit Sub
       ElseIf poly(i%).v(0) = p2% And poly(i%).v(1) = p1% And poly(i%).direction <> d Then
           polygon_no% = i%
           For j% = 2 To n% - 1
            If poly(i%).v(j%) = 0 Then
             GoTo set_polygon_mark0
            End If
           Next j%
            Exit Sub
       End If
    End If
Next i%
If last_conditions.last_cond(1).poly_no Mod 10 = 0 Then
ReDim Preserve poly(last_conditions.last_cond(1).poly_no + 10) As polygon
End If
last_conditions.last_cond(1).poly_no = last_conditions.last_cond(1).poly_no + 1
polygon_no% = last_conditions.last_cond(1).poly_no
End If
set_polygon_mark0:
tl(0) = line_number(p1%, p2%, pointapi0, pointapi0, _
                    depend_condition(point_, p1%), depend_condition(point_, p2%), _
                    condition, condition_color, 1, 0) '画第一条边
tl(1) = tl(0)
'设置基点
poly(polygon_no%).direction = d
poly(polygon_no%).is_e_polygon = is_e
'MDIForm1.Toolbar1.Buttons(21).Image = 33
Call set_polygon2(p1%, p2%, n%, poly(polygon_no%), False)
'p.v(0) = p1%
' p.coord(0) = m_poi(p1%).data(0).data0.coordinate
'p.v(1) = p2%
' p.coord(1) = m_poi(p2%).data(0).data0.coordinate

' p.total_v = n%
'****************
'计算其他顶点
'For i% = 2 To n% - 1
' last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
' MDIForm1.Toolbar1.Buttons(21).Image = 33
'  If D = False Then
'   t_coord.X = p.coord(i% - 1).X + _
'   (p.coord(i% - 2).X - p.coord(i% - 1).X) * Cos(A!) - _
'     (p.coord(i% - 2).Y - p.coord(i% - 1).Y) * Sin(A!)
'   t_coord.Y = p.coord(i% - 1).Y + _
'   (p.coord(i% - 2).X - p.coord(i% - 1).X) * Sin(A!) + _
'     (p.coord(i% - 2).Y - p.coord(i% - 1).Y) * Cos(A!)
'   Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord)
'  Else
'    t_coord.X = p.coord(i% - 1).X + _
'     (p.coord(i% - 2).X - p.coord(i% - 1).X) * Cos(A!) + _
'      (p.coord(i% - 2).Y - p.coord(i% - 1).Y) * Sin(A!)
'    t_coord.Y = p.coord(i% - 1).Y - _
'     (p.coord(i% - 2).X - p.coord(i% - 1).X) * Sin(A!) + _
'      (p.coord(i% - 2).Y - p.coord(i% - 1).Y) * Cos(A!)
'   Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord)
'End If
'  Call set_point_visible(last_conditions.last_cond(1).point_no, 1)
'  Call set_point_degree(last_conditions.last_cond(1).point_no, 0)
'  p.v(i%) = last_conditions.last_cond(1).point_no
'   p.coord(i%).X = m_poi(last_conditions.last_cond(1).point_no).data(0).data0.coordinate.X
'   p.coord(i%).Y = m_poi(last_conditions.last_cond(1).point_no).data(0).data0.coordinate.Y
'   Call get_new_char(p.v(i%))
      'Call draw_point(Draw_form, poi(p.v(i%)), 0, display)'
' Next i%
 '正四边形共线边合并
 If is_e = True Then
 If poly(polygon_no%).total_v = 4 Then
 ElseIf poly(polygon_no%).total_v = 4 Then
 For i% = 1 To m_lin(tl(0)).data(0).in_verti(0).line_no
  tp = is_line_line_intersect(tl(0), m_lin(tl(0)).data(0).in_verti(i%).line_no, 0, 0, False)
  If tp = poly(polygon_no%).v(0) Then
     record_0.data0.condition_data.condition_no = 0
     Call add_point_to_line(poly(polygon_no%).v(3), m_lin(tl(0)).data(0).in_verti(i%).line_no, 0, display, _
           True, 0)
  ElseIf tp = poly(polygon_no%).v(1) Then
     record_0.data0.condition_data.condition_no = 0
     Call add_point_to_line(poly(polygon_no%).v(2), m_lin(tl(0)).data(0).in_verti(i%).line_no, 0, display, _
           True, 0)
  End If
 Next i%
'  End If
'画四边形
Call draw_polygon4(poly(polygon_no%).v(0), _
                   poly(polygon_no%).v(1), _
                   poly(polygon_no%).v(2), _
                   poly(polygon_no%).v(3), condition)
                   
temp_record.record_.display_no = -num%
 Call set_dverti(tl(0), tl(1), temp_record, 0, 0, True)
 Call set_dverti(tl(1), tl(2), temp_record, 0, 0, True)
 Call set_dverti(tl(2), tl(3), temp_record, 0, 0, True)
 Call set_dverti(tl(0), tl(3), temp_record, 0, 0, True)
End If
End If
End Sub
Public Sub draw_polygon(p As polygon, dir%)
Dim i%
Dim v_c%, l_c%
If p.total_v = 0 Then
 Exit Sub
End If
If dir% = 1 Then
v_c% = 12
l_c% = 12
Else
v_c% = 10
l_c% = 10
End If
For i% = 0 To p.total_v - 1
'Call BPset(Draw_form, p.coord(i%).X, p.coord(i%).Y,
'     "", v_c%, display)
Call draw_plus_point(Draw_form, p.v(i%), p.coord(i%), display)
Next i%
For i% = 0 To p.total_v - 2
Draw_form.Line (p.coord(i%).X, p.coord(i%).Y)- _
 (p.coord(i% + 1).X, p.coord(i% + 1).Y), QBColor(l_c%)
Next i%
If p.total_v > 2 Then
Draw_form.Line (p.coord(0).X, p.coord(0).Y)- _
 (p.coord(p.total_v - 1).X, p.coord(p.total_v - 1).Y), QBColor(l_c%)
End If
End Sub

Public Function read_line1(p1 As POINTAPI, p2 As POINTAPI, _
   in_coord As POINTAPI, out_coord As POINTAPI, out_p%, t!, jud_dis As Integer, is_change As Boolean) As Boolean
Dim r&, s&
 'out_coord = in_coord
 If Abs(p2.X - p1.X) > 10000 Or Abs(p2.Y - p1.Y) > 10000 Then
  read_line1 = False
   Exit Function
 End If
 r& = (p2.X - p1.X) ^ 2 + (p2.Y - p1.Y) ^ 2
 If r& = 0 Then
  read_line1 = False
   Exit Function
 End If
 s& = (p2.X - p1.X) * (in_coord.X - p1.X) + _
       (p2.Y - p1.Y) * (in_coord.Y - p1.Y)
    t! = s& / r&
      out_coord.X = p1.X + (p2.X - p1.X) * t!
      out_coord.Y = p1.Y + (p2.Y - p1.Y) * t!
       If Abs(out_coord.X - in_coord.X) < jud_dis And Abs(out_coord.Y - in_coord.Y) < jud_dis Then
        read_line1 = True
         If out_p% > 0 Then
          Call set_point_coordinate(out_p%, out_coord, is_change)
         End If
       End If
End Function

Public Function read_line_33(p1 As POINTAPI, p2 As POINTAPI, _
   p3 As POINTAPI, in_coord As POINTAPI, out_coord As POINTAPI, out_p%, t!, is_change As Boolean) As Boolean
Dim r&, s&
 r& = (p3.X - p2.X) ^ 2 + (p3.Y - p2.Y) ^ 2
 s& = (p3.Y - p2.Y) * (in_coord.X - p2.X) - _
       (p3.X - p2.X) * (in_coord.Y - p2.Y)
    
    t! = s& / r&
      out_coord.X = p1.X + (p3.Y - p2.Y) * t!
      out_coord.Y = p1.Y - (p3.X - p2.X) * t!
       If Abs(out_coord.X - in_coord.X) > 5 And Abs(out_coord.Y - in_coord.Y) > 5 Then
        read_line_33 = True
       End If
       If out_p% > 0 Then
       Call set_point_coordinate(out_p%, out_coord, is_change)
       End If
End Function


Public Function set_polygon1(ByVal p1%, ByVal p2%, _
    n%, p As polygon, ByVal d As Boolean, no%, _
     ByVal no_reduce As Byte) As Byte
Dim i%, l%, j%
Dim A!
Dim v$
p.direction = d
p.is_e_polygon = True
Call set_polygon2(p1%, p2%, n, p, False)
If p.total_v = 4 Then
   Call draw_polygon4(p.v(0), p.v(1), p.v(2), p.v(3), condition)
Else
   Call draw_triangle(p.v(0), p.v(1), p.v(2), condition)
End If
If p.total_v = 4 Then
For i% = 1 To last_conditions.last_cond(1).point_no
 If i% <> p.v(0) And i% <> p.v(1) Then
  If is_dverti0(line_number0(i%, p.v(0), 0, 0), l%) Then
   record_0.data0.condition_data.condition_no = 0
   Call add_point_to_line(p.v(3), line_number0(p.v(0), i%, 0, 0), 0, _
                          display, True, 0)
  ElseIf is_dverti0(line_number0(i%, p.v(1), 0, 0), l%) Then
   record_0.data0.condition_data.condition_no = 0
   Call add_point_to_line(p.v(2), line_number0(p.v(1), i%, 0, 0), 0, _
                          display, True, 0)
  End If
 End If
Next i%
End If

End Function
Public Sub set_polygon2(ByVal p1%, ByVal p2%, _
  ByVal n%, p As polygon, is_change As Boolean)
'无图
Dim i%, l%, j%
Dim A!
p.total_v = n%
A! = PI * (n% - 2) / n%
   If p1% > 0 And p2% > 0 And p.v(0) = 0 And p.v(1) = 0 Then
      p.v(0) = p1%
      p.v(1) = p2%
   End If
   p.coord(0) = m_poi(p.v(0)).data(0).data0.coordinate
   p.coord(1) = m_poi(p.v(1)).data(0).data0.coordinate
For i% = 2 To n% - 1
  If p.direction = False Then
   t_coord.X = p.coord(i% - 1).X + _
    (p.coord(i% - 2).X - p.coord(i% - 1).X) * Cos(A!) - _
     (p.coord(i% - 2).Y - p.coord(i% - 1).Y) * Sin(A!)
   t_coord.Y = p.coord(i% - 1).Y + _
    (p.coord(i% - 2).X - p.coord(i% - 1).X) * Sin(A!) + _
     (p.coord(i% - 2).Y - p.coord(i% - 1).Y) * Cos(A!)
  Else
   t_coord.X = p.coord(i% - 1).X + _
    (p.coord(i% - 2).X - p.coord(i% - 1).X) * Cos(A!) + _
     (p.coord(i% - 2).Y - p.coord(i% - 1).Y) * Sin(A!)
   t_coord.Y = p.coord(i% - 1).Y - _
   (p.coord(i% - 2).X - p.coord(i% - 1).X) * Sin(A!) + _
     (p.coord(i% - 2).Y - p.coord(i% - 1).Y) * Cos(A!)
  End If
   p.coord(i%) = t_coord
   Call draw_new_point(p.coord(i%), depend_condition(point_, p1%), depend_condition(point_, p2%), _
          False, True, 0)
   p.v(i%) = temp_point(draw_step).no
   draw_step = draw_step + 1
   Call set_point_coordinate(p.v(i%), p.coord(i%), is_change)
   p.line_no(i% - 1) = set_line(p.v(i% - 1), p.v(i%), 0, 0, pointapi0, pointapi0, pointapi0, _
           depend_condition(point_, p.v(i% - 1)), depend_condition(point_, p.v(i%)), condition, condition_color, 1, 0)
   Call set_parent(point_, p1%, point_, p.v(i%), 0)
   Call set_parent(point_, p2%, point_, p.v(i%), 0)
  ' m_poi(p.v(i%)).data(0).degree = 0
   'Call get_new_char(p.v(i%))
   'Call set_point_visible(p.v(i%), 1, False)
Next i%
   p.line_no(n%) = set_line(p.v(n% - 1), p.v(0), 0, 0, pointapi0, pointapi0, pointapi0, _
    depend_condition(point_, p.v(n% - 1)), depend_condition(point_, p.v(0)), condition, condition_color, 1, 0)
               Call set_wenti_cond_16_12_9_8(list_type_for_draw, temp_point(0).no, temp_point(1).no, _
                   temp_point(2).no, temp_point(3).no, temp_point(4).no, temp_point(5).no)
End Sub

Public Sub set_tangent_line_for_two_circle(ByVal c1%, ByVal c2%, _
                 ByVal no_reduce As Byte, Optional tangent_line_no As Integer = 0, _
                   Optional tangent_line_ty As Integer = 0) ', out_c1%, out_c2%)
Dim D1&, D2&, r2&, tD&
Dim sr&
Dim co!, si!
Dim is_tangent_line_change As Boolean
Dim p_coord(3) As POINTAPI
Dim p As POINTAPI
If no_reduce = 255 Then
 Exit Sub
End If
If c1% > c2% Then
Call exchange_two_integer(c1, c2)
End If
'r& = (m_Circ(out_c1%).data(0).data0.c_coord.X - _
       m_Circ(out_c2%).data(0).data0.c_coord.X) ^ 2 + _
     (m_Circ(out_c1%).data(0).data0.c_coord.Y - _
       m_Circ(c2%).data(0).data0.c_coord.Y) ^ 2
        sr& = distance_of_two_POINTAPI(m_Circ(c1%).data(0).data0.c_coord, m_Circ(c2%).data(0).data0.c_coord) '圆心距
        ' temp_k1! = CSng((m_Circ(c1%).data(0).data0.c_coord.X - _
                     m_Circ(c2%).data(0).data0.c_coord.X) / sr&) 'cos
        ' temp_k2! = CSng((m_Circ(c1%).data(0).data0.c_coord.Y - _
                     m_Circ(c2%).data(0).data0.c_coord.Y) / sr&) 'sin
D1& = m_Circ(c2%).data(0).data0.radii - m_Circ(c1%).data(0).data0.radii
D2& = m_Circ(c2%).data(0).data0.radii + m_Circ(c1%).data(0).data0.radii
tD& = Abs(D1&)
'****************************************************************************************
'If tangent_line_ty = 0 Then '设置切线
If sr& > D2& Then '内公切线(圆心距比两圆的半径和大)
Call inter_point_circle_circle_by_pointapi(m_Circ(c1%).data(0).data0.c_coord, D2&, _
                  mid_POINTAPI(m_Circ(c1%).data(0).data0.c_coord, m_Circ(c2%).data(0).data0.c_coord), _
                   sr& / 2, p_coord(0), p_coord(2))
If tangent_line_ty = 0 Or tangent_line_ty = inner_tangent_line_by_two_circle12 Then
         p = minus_POINTAPI(p_coord(0), m_Circ(c1%).data(0).data0.c_coord)
         D2& = abs_POINTAPI(p)
     p_coord(0) = add_POINTAPI(m_Circ(c1%).data(0).data0.c_coord, _
               time_POINTAPI_by_number(p, m_Circ(c1%).data(0).data0.radii / D2&))
     p_coord(1) = minus_POINTAPI(m_Circ(c2%).data(0).data0.c_coord, _
               time_POINTAPI_by_number(p, m_Circ(c2%).data(0).data0.radii / D2&))
     If tangent_line_ty = 0 Then
          Call set_tangent_line_data(p_coord(0), p_coord(1), c1%, c2%, _
                depend_condition(circle_, c1%), depend_condition(circle_, c2%), 1, inner_tangent_line_by_two_circle12)
     Else
        is_tangent_line_change = True
     End If
End If
'**********************************************************************************************************
If tangent_line_ty = 0 Or tangent_line_ty = inner_tangent_line_by_two_circle21 Then
         p = minus_POINTAPI(p_coord(2), m_Circ(c1%).data(0).data0.c_coord)
         D2& = distance_of_two_POINTAPI(p_coord(2), m_Circ(c1%).data(0).data0.c_coord)
  p_coord(0) = add_POINTAPI(m_Circ(c1%).data(0).data0.c_coord, _
               time_POINTAPI_by_number(p, m_Circ(c1%).data(0).data0.radii / D2&))
  p_coord(1) = minus_POINTAPI(m_Circ(c2%).data(0).data0.c_coord, _
               time_POINTAPI_by_number(p, m_Circ(c2%).data(0).data0.radii / D2&))
        If tangent_line_ty = 0 Then
          Call set_tangent_line_data(p_coord(0), p_coord(1), c1%, c2%, _
                depend_condition(circle_, c1%), depend_condition(circle_, c2%), 1, inner_tangent_line_by_two_circle21)
        Else
           is_tangent_line_change = True
        End If
End If
'************************************************************************************************
End If
If tD& < 3 Then '等圆，外公切线
p = verti_POINTAPI(minus_POINTAPI(m_Circ(c2%).data(0).data0.c_coord, m_Circ(c1%).data(0).data0.c_coord))
D2& = abs_POINTAPI(p)
If tangent_line_ty = 0 Or tangent_line_ty = out_tangent_line_by_two_circle21 Then
p_coord(0) = add_POINTAPI(m_Circ(c1%).data(0).data0.c_coord, _
                time_POINTAPI_by_number(p, m_Circ(c1%).data(0).data0.radii / D2&))
p_coord(1) = add_POINTAPI(m_Circ(c2%).data(0).data0.c_coord, _
                time_POINTAPI_by_number(p, m_Circ(c2%).data(0).data0.radii / D2&))
       If tangent_line_ty = 0 Then
         Call set_tangent_line_data(p_coord(0), p_coord(1), c1%, c2%, _
                depend_condition(circle_, c1%), depend_condition(circle_, c2%), 1, out_tangent_line_by_two_circle21)
       Else
        is_tangent_line_change = True
       End If
End If
If tangent_line_ty = 0 Or tangent_line_ty = out_tangent_line_by_two_circle12 Then
p_coord(0) = minus_POINTAPI(m_Circ(c1%).data(0).data0.c_coord, _
                time_POINTAPI_by_number(p, m_Circ(c1%).data(0).data0.radii / D2&))
p_coord(1) = minus_POINTAPI(m_Circ(c2%).data(0).data0.c_coord, _
                time_POINTAPI_by_number(p, m_Circ(c2%).data(0).data0.radii / D2&))
       If tangent_line_ty = 0 Then
            Call set_tangent_line_data(p_coord(0), p_coord(1), c1%, c2%, _
                depend_condition(circle_, c1%), depend_condition(circle_, c2%), 1, out_tangent_line_by_two_circle12)
       Else
           is_tangent_line_change = True
       End If
   End If
ElseIf sr& > tD& Then  '连心线长大于两圆半径差，有外切公切线
   'D1& = Abs(D1&)
    Call inter_point_circle_circle_by_pointapi(mid_POINTAPI(m_Circ(c1%).data(0).data0.c_coord, _
                   m_Circ(c2%).data(0).data0.c_coord), sr& / 2, _
                      m_Circ(c2%).data(0).data0.c_coord, tD&, p_coord(0), p_coord(2))
      If D1 < 0 Then
       p_coord(3) = p_coord(2)
       p_coord(2) = p_coord(0)
       p_coord(0) = p_coord(3)
      End If
   If tangent_line_ty = 0 Or tangent_line_ty = out_tangent_line_by_two_circle12 Then
    p = minus_POINTAPI(p_coord(0), m_Circ(c2%).data(0).data0.c_coord)
    tD& = abs_POINTAPI(p)
     If D1& < 0 Then
        tD& = -tD&
     End If
    p_coord(0) = add_POINTAPI(m_Circ(c1%).data(0).data0.c_coord, _
                time_POINTAPI_by_number(p, m_Circ(c1%).data(0).data0.radii / tD&))
    p_coord(1) = add_POINTAPI(m_Circ(c2%).data(0).data0.c_coord, _
                time_POINTAPI_by_number(p, m_Circ(c2%).data(0).data0.radii / tD&))
        If tangent_line_ty = 0 Then
          Call set_tangent_line_data(p_coord(0), p_coord(1), c1%, c2%, _
                depend_condition(circle_, c2%), depend_condition(circle_, c1%), 1, out_tangent_line_by_two_circle12)
        Else
         is_tangent_line_change = True
        End If
  End If
'*********************************************************************************************************
 If tangent_line_ty = 0 Or tangent_line_ty = out_tangent_line_by_two_circle21 Then
    p = minus_POINTAPI(p_coord(2), m_Circ(c2%).data(0).data0.c_coord)
    tD& = abs_POINTAPI(p)
     If D1& < 0 Then
      tD& = -tD&
     End If
    p_coord(0) = add_POINTAPI(m_Circ(c1%).data(0).data0.c_coord, _
                time_POINTAPI_by_number(p, m_Circ(c1%).data(0).data0.radii / tD&))
    p_coord(1) = add_POINTAPI(m_Circ(c2%).data(0).data0.c_coord, _
                time_POINTAPI_by_number(p, m_Circ(c2%).data(0).data0.radii / tD&))
        If tangent_line_ty = 0 Then
          Call set_tangent_line_data(p_coord(0), p_coord(1), c1%, c2%, _
                depend_condition(circle_, c2%), depend_condition(circle_, c1%), 1, out_tangent_line_by_two_circle21)
        Else
                 is_tangent_line_change = True
        End If
 End If
 '************************************************************************************************************
 End If
 
If is_tangent_line_change Then
        tangent_line(tangent_line_no%).data(0).coordinate(0) = p_coord(0)
        tangent_line(tangent_line_no%).data(0).coordinate(1) = p_coord(1)
        m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).data0.coordinate = p_coord(0)
        m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).data0.coordinate = p_coord(1)
        'If m_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).is_change = False Then
        'm_poi(tangent_line(tangent_line_no%).data(0).poi(0)).data(0).is_change = True
        Call change_m_point(tangent_line(tangent_line_no%).data(0).poi(0))
        'End If
        'If m_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).is_change = False Then
        'm_poi(tangent_line(tangent_line_no%).data(0).poi(1)).data(0).is_change = True
        Call change_m_point(tangent_line(tangent_line_no%).data(0).poi(1))
        'End If
End If
End Sub
Public Function read_circle_from_three_point(p1%, p2%, p3%) As Integer
Dim i%, j%, k%
For i% = 1 To last_conditions.last_cond(1).circle_no
 If m_Circ(i%).data(0).data0.in_point(0) >= 3 Then
  k% = 0
  For j% = 1 To m_Circ(i%).data(0).data0.in_point(0)
  If m_Circ(i%).data(0).data0.in_point(j%) = p1% Or m_Circ(i%).data(0).data0.in_point(j%) = p2% _
     Or m_Circ(i%).data(0).data0.in_point(j%) = p3% Then
      k% = k% + 1
       If k% = 3 Then
        read_circle_from_three_point = i%
         Exit Function
       End If
  End If
  Next j%
 End If
Next i%
End Function
Public Sub set_no_display_point_on_line(ByVal l%, ByVal m_point%)
If m_lin(l%).data(0).data0.poi(0) = m_point% Then
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'MDIForm1.Toolbar1.Buttons(21).Image = 33
Call set_point_visible(last_conditions.last_cond(1).point_no, 0, False)
t_coord.X = 2 * m_poi(m_point%).data(0).data0.coordinate.X - _
      m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.X
t_coord.Y = 2 * m_poi(m_point%).data(0).data0.coordinate.Y - _
      m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate.Y
Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
Call get_new_char(last_conditions.last_cond(1).point_no)
record_0.data0.condition_data.condition_no = 0
Call add_point_to_line(last_conditions.last_cond(1).point_no, l%, 0, no_display, _
     True, 0)
ElseIf m_lin(l%).data(0).data0.poi(1) = m_point% Then
last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
'MDIForm1.Toolbar1.Buttons(21).Image = 33
Call set_point_visible(last_conditions.last_cond(1).point_no, 0, False)
t_coord.X = 2 * m_poi(m_point%).data(0).data0.coordinate.X - _
      m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.X
t_coord.Y = 2 * m_poi(m_point%).data(0).data0.coordinate.Y - _
      m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate.Y
Call set_point_coordinate(last_conditions.last_cond(1).point_no, t_coord, False)
Call get_new_char(last_conditions.last_cond(1).point_no)
record_0.data0.condition_data.condition_no = 0
Call add_point_to_line(last_conditions.last_cond(1).point_no, l%, 0, no_display, _
     True, 0)
End If
End Sub
Public Function read_circle0(c_data0 As circle_data0_type, _
                   input_coord As POINTAPI, out_coord As POINTAPI) As Boolean
Dim s&
s& = abs_POINTAPI(minus_POINTAPI(c_data0.c_coord, input_coord))
'(x!,y!)离圆心的距离
If Abs(s& - c_data0.radii) < 5 Then  '接近圆的半径
 out_coord.X = c_data0.c_coord.X + c_data0.radii * (input_coord.X - c_data0.c_coord.X) / s&
 out_coord.Y = c_data0.c_coord.Y + c_data0.radii * (input_coord.Y - c_data0.c_coord.Y) / s&
   read_circle0 = True
    Exit Function
End If
End Function

Public Sub set_tangent_circles(c1 As circle_data0_type, _
                p1 As POINTAPI, p2 As POINTAPI, _
                 r1&, r2&, c_d1 As POINTAPI, _
                  c_d2 As POINTAPI, t_p1 As POINTAPI, t_p1_%, _
                   t_p2 As POINTAPI, t_p2_%, ty As Byte, is_change As Boolean)
Dim r&, s&
Dim A!, b!, c!, d!, c_r_2!
Dim t1!, t2!, u!, w!, s_2!, r_2!
Dim p0 As POINTAPI
Dim tp(1) As POINTAPI
'On Error GoTo set_tangent_circles_error
If compare_two_point(p1, p2, 0, 0, 0) = -1 Then
 tp(0) = p2
 tp(1) = p1
Else
 tp(0) = p1
 tp(1) = p2
End If
If tp(1).X > -30000 And tp(1).Y > -30000 And _
   tp(1).X < 30000 And tp(1).Y < 30000 Then
c_r_2! = c1.radii ^ 2
r_2! = (tp(0).X - tp(1).X) ^ 2 + (tp(0).Y - tp(1).Y) ^ 2
s_2! = (c1.c_coord.X - tp(1).X) ^ 2 + (c1.c_coord.Y - tp(1).Y) ^ 2
u! = (tp(0).X - tp(1).X) * (tp(1).X - c1.c_coord.X) + _
     (tp(0).Y - tp(1).Y) * (tp(1).Y - c1.c_coord.Y)
w! = (tp(0).Y - tp(1).Y) * (tp(1).X - c1.c_coord.X) - _
     (tp(0).X - tp(1).X) * (tp(1).Y - c1.c_coord.Y)
A! = 4 * (w! ^ 2 - c_r_2! * r_2!)
b! = CSng(4 * w!) * (s_2! + u! - c_r_2!)
c! = CSng(s_2! + u! - c1.radii ^ 2) ^ 2 - CSng(c_r_2! * r_2!)
d! = b! ^ 2 - 4 * A! * c!
If d! > 0 Then
d! = sqr(d!)
t1! = (-b + d!) / 2 / A!
t2! = (-b - d!) / 2 / A!
p0.X = (tp(0).X + tp(1).X) / 2
p0.Y = (tp(0).Y + tp(1).Y) / 2
c_d1.X = p0.X + CLng(t1! * (tp(0).Y - tp(1).Y))
c_d1.Y = p0.Y - CLng(t1! * (tp(0).X - tp(1).X))
c_d2.X = p0.X + CLng(t2! * (tp(0).Y - tp(1).Y))
c_d2.Y = p0.Y - CLng(t2! * (tp(0).X - tp(1).X))
r1& = sqr((tp(0).X - c_d1.X) ^ 2 + (tp(0).Y - c_d1.Y) ^ 2)
r2& = sqr((tp(0).X - c_d2.X) ^ 2 + (tp(0).Y - c_d2.Y) ^ 2)
Else
Exit Sub
End If
Else
c_d1 = p1
c_d2 = p1
r& = sqr((c1.c_coord.X - tp(0).X) ^ 2 + (c1.c_coord.Y - tp(0).Y) ^ 2)
If r& > c1.radii Then
  r1& = r& - c1.radii
  r2& = r& + c1.radii
Else
  r1& = c1.radii - r&
  r2& = c1.radii + r&
End If
End If
If ty = 1 Then
r& = sqr((c1.c_coord.X - c_d1.X) ^ 2 + (c1.c_coord.Y - c_d1.Y) ^ 2)
If r& > c1.radii And r& > r1& Then
   r& = c1.radii + r1&
t_p1.X = (c1.c_coord.X * r1& + c_d1.X * c1.radii) / r&
t_p1.Y = (c1.c_coord.Y * r1& + c_d1.Y * c1.radii) / r&
Else
 If c1.radii < r1& Then
  r& = r1& - c1.radii
  t_p1.X = (c1.c_coord.X * r1& - c_d1.X * c1.radii) / r&
  t_p1.Y = (c1.c_coord.Y * r1& - c_d1.Y * c1.radii) / r&
 Else
  r& = c1.radii - r1&
  t_p1.X = (-c1.c_coord.X * r1& + c_d1.X * c1.radii) / r&
  t_p1.Y = (-c1.c_coord.Y * r1& + c_d1.Y * c1.radii) / r&
 End If
End If
ElseIf ty = 2 Then
r& = sqr((c1.c_coord.X - c_d2.X) ^ 2 + (c1.c_coord.Y - c_d2.Y) ^ 2)
If r& > c1.radii And r& > r2& Then
r& = c1.radii + r2&
t_p2.X = (c1.c_coord.X * r2& + c_d2.X * c1.radii) / r&
t_p2.Y = (c1.c_coord.Y * r2& + c_d2.Y * c1.radii) / r&
Else
 If c1.radii < r2& Then
  r& = r2& - c1.radii
  t_p2.X = (c1.c_coord.X * r2& - c_d2.X * c1.radii) / r&
  t_p2.Y = (c1.c_coord.Y * r2& - c_d2.Y * c1.radii) / r&
 Else
  r& = c1.radii - r2&
  t_p2.X = (-c1.c_coord.X * r2& + c_d2.X * c1.radii) / r&
  t_p2.Y = (-c1.c_coord.Y * r2& + c_d2.Y * c1.radii) / r&
 End If
End If
End If
If t_p1_% > 0 Then
Call set_point_coordinate(t_p1_%, t_p1, is_change)
End If
If t_p2_% > 0 Then
Call set_point_coordinate(t_p2_%, t_p2, is_change)
End If

set_tangent_circles_error:
End Sub

Public Sub draw_temp_line_for_move(point_type%, ele1%, ele2%) '以m_poi(0)传递数据, p1 As POINTAPI)
    If m_poi(0).data(0).data0.coordinate.X = 10000 And _
              m_poi(0).data(0).data0.coordinate.Y = 10000 Then
       Exit Sub
    End If
     Call draw_tangent_linee(0)
     Call draw_tangent_linee(1)
      Draw_form.Timer1.Enabled = True
     If point_type% = new_point_on_line Then
        Call set_temp_line(0, ele1%, m_poi(0).data(0).data0.coordinate)
        Call draw_tangent_linee(0)
     ElseIf point_type% = interset_point_line_line Then
        Call set_temp_line(0, ele1%, m_poi(0).data(0).data0.coordinate)
        Call set_temp_line(1, ele2%, m_poi(0).data(0).data0.coordinate)
        Call draw_tangent_linee(0)
        Call draw_tangent_linee(1)
     ElseIf point_type% = new_point_on_line_circle Then
        Call set_temp_line(0, ele1%, m_poi(0).data(0).data0.coordinate)
        Call draw_tangent_linee(0)
     End If

End Sub
Public Sub re_name_all_point()
Dim i%
 run_type = 10
 For i% = 1 To last_conditions.last_cond(1).point_no
  Call re_name_point(i%)
   Do
   DoEvents
   Loop Until re_name_ty = 0
 Next i%
 operator = ""
list_type_for_draw = 0
yidian_stop = True
 run_type = 0
End Sub
Public Sub draw_view_point_move()
Dim i%
For i% = 1 To last_conditions.last_cond(1).last_view_point_no
 Draw_form.Picture2.Line (view_point(i%).old_coordinate.X, _
           view_point(i%).old_coordinate.Y)- _
            (m_poi(view_point(i%).poi).data(0).data0.coordinate.X, _
              m_poi(view_point(i%).poi).data(0).data0.coordinate.Y), QBColor(12)
  view_point(i%).old_coordinate = m_poi(view_point(i%).poi).data(0).data0.coordinate
Next i%
'Draw_form.Picture1.Cls
Call BitBlt(Draw_form.Picture1.hdc, 0, 0, Draw_form.Picture1.width, _
       Draw_form.Picture1.Height, Draw_form.hdc, 0, 0, vbSrcCopy)
Call BitBlt(Draw_form.Picture1.hdc, 0, 0, Draw_form.Picture1.width, _
       Draw_form.Picture1.Height, Draw_form.Picture2.hdc, 0, 0, vbSrcCopy)
End Sub
Public Sub draw_arrow(ob As Object, p2_x As Long, p2_y As Long, p1_x As Long, p1_y As Long, color As Long)
Dim r!
Dim s1(1) As Single
Dim S2(1) As Single
Dim s3(1) As Single
Dim tp(1) As POINTAPI
'If regist_data.run_type = 1 Then
r! = sqr((p1_x - p2_x) ^ 2 + (p1_y - p2_y) ^ 2)
If r! > 5 Then
s1(0) = (p2_x - p1_x) / r!
s1(1) = (p2_y - p1_y) / r!
S2(0) = s1(0) * 0.92 - s1(1) * 0.4
S2(1) = s1(1) * 0.92 + s1(0) * 0.4
s3(0) = s1(0) * 0.92 + s1(1) * 0.4
s3(1) = s1(1) * 0.92 - s1(0) * 0.4
tp(0).X = p1_x + S2(0) * 15
tp(0).Y = p1_y + S2(1) * 15
tp(1).X = p1_x + s3(0) * 15
tp(1).Y = p1_y + s3(1) * 15
ob.Line (p1_x, p1_y)-(tp(0).X, tp(0).Y), color
ob.Line (p1_x, p1_y)-(tp(1).X, tp(1).Y), color
End If
'End If
End Sub

Public Sub draw_verti_mark(tA%)
If T_angle(tA%).data(0).value = "90" Then
 If T_angle(tA%).data(0).is_draw_verti_mark = False Then
     If T_angle(tA%).data(0).is_used_no = 0 Then
      Call draw_verti_mark0(T_angle(tA%).data(0).inter_point, _
        T_angle(tA%).data(0).line_no(0), T_angle(tA%).data(0).line_no(1), 1, 1)
     ElseIf T_angle(tA%).data(0).is_used_no = 1 Then
      Call draw_verti_mark0(T_angle(tA%).data(0).inter_point, _
        T_angle(tA%).data(0).line_no(0), T_angle(tA%).data(0).line_no(1), 1, 0)
     ElseIf T_angle(tA%).data(0).is_used_no = 2 Then
      Call draw_verti_mark0(T_angle(tA%).data(0).inter_point, _
        T_angle(tA%).data(0).line_no(0), T_angle(tA%).data(0).line_no(1), 0, 0)
     ElseIf T_angle(tA%).data(0).is_used_no = 3 Then
      Call draw_verti_mark0(T_angle(tA%).data(0).inter_point, _
        T_angle(tA%).data(0).line_no(0), T_angle(tA%).data(0).line_no(1), 0, 1)
     End If
     T_angle(tA%).data(0).is_draw_verti_mark = True
 End If
End If
End Sub
Sub draw_verti_mark0(ByVal p%, ByVal l1%, ByVal l2%, ByVal D1%, ByVal D2%)
 Dim tp(2) As POINTAPI
 Dim l(1) As Integer
 Dim temp_width As Integer
 If m_lin(l1%).data(0).data0.visible > 0 And m_lin(l2%).data(0).data0.visible > 0 And p% > 0 Then
 l(0) = sqr((m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate.X - _
            m_poi(m_lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate.X) ^ 2 + _
            (m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate.Y - _
            m_poi(m_lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate.Y) ^ 2)
 l(1) = sqr((m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate.X - _
            m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate.X) ^ 2 + _
            (m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate.Y - _
            m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate.Y) ^ 2)
 If D1% = 1 Then
  tp(0).X = m_poi(p%).data(0).data0.coordinate.X + _
         (m_poi(m_lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate.X - _
            m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate.X) * verti_mark_meas / l(0)
  tp(0).Y = m_poi(p%).data(0).data0.coordinate.Y + _
         (m_poi(m_lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate.Y - _
            m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate.Y) * verti_mark_meas / l(0)
 Else
  tp(0).X = m_poi(p%).data(0).data0.coordinate.X + _
         (m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate.X - _
            m_poi(m_lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate.X) * verti_mark_meas / l(0)
  tp(0).Y = m_poi(p%).data(0).data0.coordinate.Y + _
         (m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate.Y - _
            m_poi(m_lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate.Y) * verti_mark_meas / l(0)
 End If
 If D2% = 1 Then
  tp(1).X = m_poi(p%).data(0).data0.coordinate.X + _
         (m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate.X - _
            m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate.X) * verti_mark_meas / l(1)
  tp(1).Y = m_poi(p%).data(0).data0.coordinate.Y + _
         (m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate.Y - _
            m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate.Y) * verti_mark_meas / l(1)
 Else
  tp(1).X = m_poi(p%).data(0).data0.coordinate.X + _
         (m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate.X - _
            m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate.X) * verti_mark_meas / l(1)
  tp(1).Y = m_poi(p%).data(0).data0.coordinate.Y + _
         (m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate.Y - _
            m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate.Y) * verti_mark_meas / l(1)
 End If
 tp(2).X = tp(1).X + (tp(0).X - m_poi(p%).data(0).data0.coordinate.X)
 tp(2).Y = tp(1).Y + (tp(0).Y - m_poi(p%).data(0).data0.coordinate.Y)
 temp_width = Draw_form.DrawWidth
 Draw_form.DrawWidth = 1
 Draw_form.Line (tp(0).X, tp(0).Y)-(tp(2).X, tp(2).Y), QBColor(12)
 Draw_form.Line (tp(1).X, tp(1).Y)-(tp(2).X, tp(2).Y), QBColor(12)
 Draw_form.DrawWidth = temp_width
 End If
End Sub
Public Function inter_point_verti_mid_line_line(p1 As POINTAPI, p2 As POINTAPI, _
               l As line_data0_type, out_coord As POINTAPI) As Boolean
Dim p0 As point_data_type
Dim A(3) As POINTAPI
Dim b(3) As Single
Dim t As Single
A(0) = mid_POINTAPI(p1, p2)
A(1) = minus_POINTAPI(p1, p2)
A(2) = minus_POINTAPI(m_poi(l.poi(0)).data(0).data0.coordinate, _
        m_poi(l.poi(1)).data(0).data0.coordinate)
b(0) = time_POINTAPI(A(1), A(2))
b(1) = cross_time_POINTAPI(A(2), A(0))
b(1) = b(1) + cross_time_POINTAPI(m_poi(l.poi(1)).data(0).data0.coordinate, _
                    m_poi(l.poi(0)).data(0).data0.coordinate)
If b(0) <> 0 Then
 inter_point_verti_mid_line_line = True
 t = b(1) / b(0)
 out_coord.X = A(0).X + t * A(1).Y
 out_coord.Y = A(0).Y - t * A(1).X
End If
End Function
Public Function inter_point_verti_mid_line_point(p1 As POINTAPI, p2 As POINTAPI, _
               p3 As POINTAPI, out_coord As POINTAPI) As Boolean
Dim p0 As point_data_type
Dim A(1) As POINTAPI
Dim b(3) As Long
Dim t As Single
A(0) = mid_POINTAPI(p1, p2)
A(1) = minus_POINTAPI(p1, p2)
b(0) = time_POINTAPI(A(1), A(1))
b(1) = cross_time_POINTAPI(A(1), A(0))
b(1) = b(1) + cross_time_POINTAPI(p3, _
                    A(1))
If b(0) <> 0 Then
 inter_point_verti_mid_line_point = True
 t = b(1) / b(0)
 out_coord.X = A(0).X + t * A(1).Y
 out_coord.Y = A(0).Y - t * A(1).X
End If
End Function
Public Function inter_point_verti_mid_line_circle(p1 As POINTAPI, p2 As POINTAPI, c As circle_data0_type, _
        out_coord1 As POINTAPI, out_coord2 As POINTAPI) As Boolean
Dim p0 As point_data_type
Dim A(3) As POINTAPI
Dim b(3) As Single
Dim d As Single
Dim t(1) As Single
A(0) = mid_POINTAPI(p1, p2)
A(1) = minus_POINTAPI(p1, p2)
A(2) = minus_POINTAPI(A(0), c.c_coord)
b(0) = time_POINTAPI(A(1), A(1))
b(1) = cross_time_POINTAPI(A(1), A(2)) '-b/2
b(2) = time_POINTAPI(A(2), A(2)) - c.radii * c.radii
d = b(1) ^ 2 - b(0) * b(2)
If d >= 0 And b(0) <> 0 Then
   inter_point_verti_mid_line_circle = True
   d = sqr(d)
   t(0) = (b(1) + d) / b(0)
   t(1) = (b(1) - d) / b(0)
   out_coord1.X = A(0).X + t(0) * A(1).Y
   out_coord1.Y = A(0).Y - t(0) * A(1).X
   out_coord2.X = A(0).X + t(1) * A(1).Y
   out_coord2.Y = A(0).Y - t(1) * A(1).X
End If
End Function
Public Function set_point_from_aid_point(ByVal aid_point_no%) As Integer
If aid_point_no% < 90 Then
   set_point_from_aid_point = aid_point_no%
Else
End If
End Function

Public Sub plane_geometry_draw_mouse_down(Button As Integer, Shift As Integer, X As Single, Y As Single)
'平面几何作图,鼠标单击
Dim i%, k%, l%, r!, tp1%, tp2%, tl%, poly_no%
Dim linespoint(1) As Integer
Dim ty1 As Integer
Dim tA As Integer
Dim temp_s$
Dim LastInput0No%, lastInput0ConditionNo%
Dim f, ope As Integer
Dim coord As POINTAPI
Dim p_c As POINTAPI
Dim ty As Boolean
Dim t_c As circle_data_type
Dim inter_point_type%
Dim dis&
Dim p1 As POINTAPI
Dim p2 As POINTAPI
'先输先处理
Up_Enabled = False
Move_Enabled = False
draw_operate = True
Move_statue = 0
draw_step = draw_step + 1 '计算输入步数
If input_text_statue Then '文本输入状态
 Exit Sub
End If
draw_statue = True '设定图形输入
mouse_down_coord.X = Int(X)
mouse_down_coord.Y = Int(Y)
If event_statue = wait_for_modify_char Or _
     event_statue = wait_for_input_char Or _
        event_statue = input_char_again Or _
         draw_statue = False Then
'修改点的名,设置点闪烁
      If draw_new_point(mouse_down_coord, ele1, ele2, blue, False, 1) = exist_point Then    '读出已知点
          '输入点的名称
          Call C_display_wenti.m_input_char(Wenti_form.Picture1, m_poi(temp_point(draw_step).no).data(0).data0.name) '读出点的名称
          'draw_point_no = 0
      End If
Else
 '画点线圆
  init_p.X = Int(X)
   init_p.Y = Int(Y)
 '读出坐标
 mouse_move_coord.X = Int(X)
  mouse_move_coord.Y = Int(Y)
'************************************************
If Button = 2 Then '按右键
 Call display_sub_menu(X, Y) '显示子菜单
'*******************
'不可step=1用留作拖线用
ElseIf Button = 1 Then '按左键
If event_statue = wait_for_draw_point Then
  mouse_type% = 1
   event_statue = draw_point_down
End If
Select Case operator
  '读出点和点的坐标
Case "draw_point_and_line" '画点和直线
If list_type_for_draw = 0 Then
     Exit Sub
'********************************************************************
ElseIf list_type_for_draw = 1 Then
  If event_statue = ready Then
  If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 0) > 0 Then    '
      move_init = draw_step
      Up_Enabled = True '拖动画直线
      Move_Enabled = True
  End If
  End If
ElseIf list_type_for_draw = 2 Or list_type_for_draw = 3 Or list_type_for_draw = 4 Then
     '画中点画定长等长线段
        If draw_step = 0 Then
         If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 0) > 0 Then    '
            move_init = draw_step
            Up_Enabled = True
            Move_Enabled = True
         End If
       ElseIf draw_step = 1 Then
         mouse_up_no_enabled = True
       End If
'***********************************
ElseIf list_type_for_draw = 5 Then '画等长线段
      If draw_step = 0 Then
       If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 0) > 0 Then
        move_init = draw_step
        Up_Enabled = True
        Move_Enabled = True
       End If
      ElseIf draw_step = 2 Then
        If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 0, True) > 0 Then
          move_init = draw_step
           Move_Enabled = True
           temp_circle(0) = m_circle_number(1, temp_point(2).no, pointapi0, _
                      0, 0, 0, distance_of_two_POINTAPI(m_poi(temp_point(0).no).data(0).data0.coordinate, _
                          m_poi(temp_point(1).no).data(0).data0.coordinate), temp_point(0).no, temp_point(1).no, 1, 1, wenti_cond_, _
                       fill_color, False)
           MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1975, "")
       End If
      End If
ElseIf list_type_for_draw = 6 Then '画角平分线
    If draw_step = 0 Then
       If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 0) > 0 Then  '画第一点
       move_init = draw_step
       Up_Enabled = True
       Move_Enabled = True
       End If
    ElseIf draw_step = 1 Then
      mouse_up_no_enabled = True
       ' Exit Sub
    ElseIf draw_step = 2 Then
       If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 0) > 0 Then
           If temp_point(2).no <> temp_point(0).no And temp_point(2).no <> temp_point(1).no Then  '第二条边的起点，与第一条边的两个端点都不同
               temp_point(2).no = temp_point(1).no
                draw_step = 3
'                 mouse_down_coord = m_poi(temp_point(2).no).data(0).data0.coordinate
                   mouse_up_no_enabled = True
                     Move_Enabled = False
           Else '
             If temp_point(2).no = temp_point(0).no Then '第二条边的起点，与第一条边第一个端点相同
                Call exchange_two_integer(temp_point(0).no, temp_point(1).no) '交换，保证第二条边的起点，与第一条边第第二个端点相同
             End If
              move_init = draw_step
              'Up_Enabled = True
              Move_Enabled = True
           End If
        Up_Enabled = True
        End If
    ElseIf draw_step = 3 Then
        mouse_up_no_enabled = True
    ElseIf draw_step = 4 Then
          Call set_select_point
          Call set_select_line(temp_line(2))
          Call set_select_circle
          Call set_forbid_point(temp_point(1).no)
          Call set_forbid_line
          Call set_forbid_circle
        If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 255, True) > 0 Then   '
             Call init_draw_data
             list_type_for_draw = 1
        Else
            draw_step = 2 '撤消操作
        End If

    End If
End If
'************************************************************
Case "draw_circle" 'Then                     '************
'*********************************************************
If list_type_for_draw = 0 Then
'**********************************************************************
 Exit Sub
 '****************************************************************************
ElseIf list_type_for_draw = 1 Then '画圆
'*****************************************************************************
 If draw_step = 0 Then
   If circle_with_center Then
     Up_Enabled = True
   Else
     Up_Enabled = False
   End If
 If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 0) > 0 Then
       move_init = draw_step
       Up_Enabled = True
       Move_Enabled = True
       m_temp_circle_for_input.data(0).c_coord = mouse_down_coord ' m_poi(temp_point(0).no).data(0).data0.coordinate
       m_temp_circle_for_input.data(0).color = fill_color
       m_temp_circle_for_input.is_using = True
       'temp_circle(0) = m_circle_number(1, temp_point(0).no, pointapi0, _
                         0, 0, 0, 0, 0, 0, 1, 1, condition, fill_color, True)
 End If
 End If
 '***************************************************************************
ElseIf list_type_for_draw = 2 Then '三点画无心圆
'****************************************************************************
 If draw_step = 0 Then
    If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 0) > 0 Then
       Up_Enabled = False
       Move_Enabled = False
       m_temp_circle_for_input.data(0).c_coord = mouse_down_coord 'm_poi(temp_point(0).no).data(0).data0.coordinate
       m_temp_circle_for_input.data(0).color = fill_color
       m_temp_circle_for_input.is_using = True

       'temp_circle(0) = m_circle_number(1, temp_point(0).no, pointapi0, _
                         0, 0, 0, 0, 0, 0, 1, 1, aid_condition, fill_color, True)
    End If
 ElseIf draw_step = 1 Then
         Call set_select_point '         Call set_select_line
         Call set_select_circle(temp_circle(0))
         Call set_forbid_point(temp_point(0).no)
         Call set_forbid_line
         Call set_forbid_circle
 If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 255) <= 0 Then
       draw_step = 0
 Else
    m_temp_circle_for_input.data(0).c_coord = mid_POINTAPI(m_poi(temp_point(0).no).data(0).data0.coordinate, _
                             m_poi(temp_point(1).no).data(0).data0.coordinate)
    m_temp_circle_for_input.data(0).radii = distance_of_two_POINTAPI(m_poi(temp_point(0).no).data(0).data0.coordinate, _
                             m_poi(temp_point(1).no).data(0).data0.coordinate) / 2
    m_temp_circle_for_input.data(0).in_point(0) = 2
     m_temp_circle_for_input.data(0).in_point(1) = temp_point(0).no
     m_temp_circle_for_input.data(0).in_point(2) = temp_point(1).no
    Call draw_temp_circle_for_input
 End If
ElseIf draw_step = 2 Then
       '画第三个点,
         Call set_select_point
         Call set_select_line
         Call set_select_circle(temp_circle(0))
         Call set_forbid_point(temp_point(0).no, temp_point(1).no)
         Call set_forbid_line
         Call set_forbid_circle
        m_temp_circle_for_input.is_using = True
        m_temp_circle_for_input.data(0).color = 15 'condition_color
            Call draw_temp_circle_for_input
     If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 255) > 0 Then
        'temp_circle(0) = m_circle_number(1, 0, m_temp_circle_for_input.data(0).c_coord, temp_point(0).no, _
                        temp_point(1).no, temp_point(2).no, m_temp_circle_for_input.data(0).radii, _
                         0, 0, 0, 1, 1, condition_color, True)
        Call init_draw_data
        operator = "draw_point_and_line"
         list_type_for_draw = 1
     Else
        draw_step = 1
     End If
 End If
 '*********************************************************
 ElseIf list_type_for_draw = 3 Then '点到圆的切线
 '*************************************************************
  If draw_step = 0 Then
  operat_is_acting = True
 '选圆
      Call read_three_point_circle(mouse_down_coord, 0, 3)
       mouse_up_no_enabled = True
   ElseIf draw_step = 1 Then
      Call read_three_point_circle(mouse_down_coord, 0, 3)
      mouse_up_no_enabled = True
   ElseIf draw_step = 2 Then
    Call read_three_point_circle(mouse_down_coord, 0, 3)
    mouse_up_no_enabled = True
   ElseIf draw_step = 3 Then
    If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 255) > 0 Then
     Call set_tangent_line_from_point_to_circle(temp_circle(0), temp_point(3).no, 0)
    End If
    mouse_up_no_enabled = True
   ElseIf draw_step = 4 Then
          Call set_select_point
          Call set_select_line
          Call set_select_circle
          Call set_forbid_point
          Call set_forbid_line
          Call set_forbid_circle(temp_circle(0))
    If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 5) > 0 Then
           Call init_draw_data
            operator = "draw_point_and_line"
            list_type_for_draw = 1
    Else
      draw_step = 3
    End If
     mouse_up_no_enabled = True
   End If
 '*************************************************************
 ElseIf list_type_for_draw = 4 Then '画两圆的公切线
 '*************************************************************
 '画切线
 If draw_step = 0 Then
  operat_is_acting = True
 '选圆
    Call read_three_point_circle(mouse_down_coord, 0, 4)
         mouse_up_no_enabled = True
 ElseIf draw_step = 1 Then
    Call read_three_point_circle(mouse_down_coord, 0, 4)
         mouse_up_no_enabled = True
 ElseIf draw_step = 2 Then
    Call read_three_point_circle(mouse_down_coord, 0, 4)
         mouse_up_no_enabled = True
ElseIf draw_step = 3 Or draw_step = 4 Or draw_step = 5 Then
'选切线
   If read_three_point_circle(mouse_down_coord, 1, 4) Then
    If draw_step < 6 Then
     If temp_circle(0) > temp_circle(1) Then
        Call exchange_two_integer(temp_circle(0), temp_circle(1))
     End If
      Call set_tangent_line_for_two_circle(temp_circle(0), temp_circle(1), 0)
       draw_step = 5
     End If
    End If
         mouse_up_no_enabled = True
ElseIf draw_step = 6 Then
    If draw_new_point(mouse_down_coord, ele1, ele2, blue, True, 5) > 0 Then   '点必需落在切线上
       Call init_draw_data
        operator = "draw_point_and_line"
         list_type_for_draw = 1
    Else
       draw_step = 5
    End If
         mouse_up_no_enabled = True
End If
'**********************************************************************************
ElseIf list_type_for_draw = 5 Then '圆相切于已有圆(或直线)
'***************************************************************************************8
 Draw_form.List1.visible = False
 If draw_step = 0 Then '画圆相切,选第一圆
   Call read_three_point_circle(mouse_down_coord, 0, 5)
    mouse_up_no_enabled = True
ElseIf draw_step = 1 Then
   Call read_three_point_circle(mouse_down_coord, 0, 5)
    mouse_up_no_enabled = True
ElseIf draw_step = 2 Then
   Call read_three_point_circle(mouse_down_coord, 0, 5)
   mouse_up_no_enabled = True
 ElseIf draw_step = 3 Then '画相切圆
    Call draw_new_point(mouse_down_coord, ele1, ele2, red, True, 0)
         'Exit Sub
   '    End If
     If temp_circles_for_draw(0) = 1 Then
     dis& = distance_of_two_POINTAPI(m_poi(temp_point(3).no).data(0).data0.coordinate, m_Circ(temp_circles_for_draw(1)).data(0).data0.c_coord)
      Call inter_point_line_circle3(m_poi(temp_point(3).no).data(0).data0.coordinate, paral_, _
             m_poi(temp_point(3).no).data(0).data0.coordinate, m_Circ(temp_circles_for_draw(1)).data(0).data0.c_coord, _
              m_Circ(temp_circles_for_draw(1)).data(0).data0, p1, 0, p2, 0, 0, 0)
         temp_tangent_circle_no(0) = set_tangent_circle_data(m_poi(temp_point(3).no).data(0).data0.coordinate, _
               dis& + m_Circ(temp_circles_for_draw(1)).data(0).data0.radii, _
                0, p1, depend_condition(circle_, temp_circles_for_draw(1)), 0, pointapi0, depend_condition(0, 0), new_point_on_line_circle12, _
                  temp_tangent_circle_no(0))
                   m_tangent_circle(temp_tangent_circle_no(0)).data(0).center = temp_point(3).no
         temp_tangent_circle_no(1) = set_tangent_circle_data(m_poi(temp_point(3).no).data(0).data0.coordinate, _
               Abs(dis& - m_Circ(temp_circles_for_draw(1)).data(0).data0.radii), _
                0, p2, depend_condition(circle_, temp_circles_for_draw(1)), 0, pointapi0, depend_condition(0, 0), new_point_on_line_circle21, _
                  temp_tangent_circle_no(1))
                   m_tangent_circle(temp_tangent_circle_no(1)).data(0).center = temp_point(3).no
     ElseIf temp_lines_for_draw(0) = 1 Then
     End If
     mouse_up_no_enabled = True
 '*************************************************************************
 '**********************************************************************
ElseIf draw_step = 4 Then
  Call draw_new_point(mouse_down_coord, ele1, ele2, red, True, 0, True)
  'temp_circle(6) = C_display_picture.read_m_circle(mouse_down_coord.X, mouse_down_coord.Y, aid_condition)
       Call init_draw_data
 End If
 ElseIf list_type_for_draw = 6 Then
 If draw_step = 0 Then
  operat_is_acting = True
 '选圆
    Call read_three_point_circle(mouse_down_coord, 0, 6)
         mouse_up_no_enabled = True
 ElseIf draw_step = 1 Then
    Call read_three_point_circle(mouse_down_coord, 0, 6)
         mouse_up_no_enabled = True
 ElseIf draw_step = 2 Then
    Call read_three_point_circle(mouse_down_coord, 0, 6)
         mouse_up_no_enabled = True
ElseIf draw_step = 3 Or draw_step = 4 Or draw_step = 5 Then
'选切线
   If read_three_point_circle(mouse_down_coord, 1, 6) Then
    If draw_step < 6 Then
      If temp_circle(0) > temp_circle(1) Then
        Call exchange_two_integer(temp_circle(0), temp_circle(1))
      End If
     End If
    End If
         mouse_up_no_enabled = True
ElseIf draw_step = 6 Then
  tp1% = m_point_number(m_tangent_circle(temp_tangent_circle_no(0)).data(0).data0(0).circle_center, condition, 1, _
          condition_color, "", depend_condition(circle_, temp_circle(0)), depend_condition(circle_, temp_circle(1)), 0, _
           True)
 m_tangent_circle(temp_tangent_circle_no(0)).data(0).circle_no = m_circle_number(1, tp1%, m_poi(tp1%).data(0).data0.coordinate, _
        0, 0, 0, m_tangent_circle(temp_tangent_circle_no(0)).data(0).data0(0).circle_radii, _
          0, 0, 0, 1, 0, condition_color, True)
 Call set_tangent_circle_data(m_tangent_circle(temp_tangent_circle_no(0)).data(0).data0(0).circle_center, 0, _
                  0, pointapi0, depend_condition(circle_, temp_circle(0)), 0, pointapi0, _
                   depend_condition(circle_, temp_circle(1)), 0, 0)
 Call set_wenti_cond_71_70_69(m_tangent_circle(temp_tangent_circle_no(0)).data(0).circle_no, _
                      temp_circle(0), temp_circle(1))
   '  If draw_new_point(m_tangent_circle(temp_tangent_circle_no(0)).data(0).data0(0).circle_center, ele1, ele2, blue, True, 5) > 0 Then  '点必需落在切线上
    Call init_draw_data
 End If
End If
Case "paral_verti" '画平行垂直线，垂直平分线
 '*************************************************
 'draw_step = 0 到 draw_step = 1 画标准直线
If list_type_for_draw = 1 Then
   If draw_step = 0 Then
     If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 0) > 0 Then     '
      move_init = draw_step
      Up_Enabled = True '拖动画直线
      Move_Enabled = True
      End If
   ElseIf draw_step = 1 Then
     mouse_up_no_enabled = True
   ElseIf draw_step = 2 Then
         Call set_select_point
         Call set_select_line
         Call set_select_circle
         Call set_forbid_point
         Call set_forbid_line
         Call set_forbid_circle
        If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 255, 0) > 0 Then
         Call C_display_picture.set_aid_line_start_point(temp_point(2).no, temp_line(0))
         ' Call set_temp_paral_and_vertical_line(temp_point(2).no, temp_line(0))
           MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1995, "")
        End If
        mouse_up_no_enabled = True
  ElseIf draw_step = 3 Then
         Call set_select_point
         Call set_select_line
         Call set_select_circle
         Call set_forbid_point(temp_point(2).no)
         Call set_forbid_line
         Call set_forbid_circle
      If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 255, True) > 0 Then
               Call init_draw_data
               operat_is_acting = False
      Else
        draw_step = 2
      End If
  End If
ElseIf list_type_for_draw = 2 Then
 If draw_step = 0 Then
  If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 0) > 0 Then     '
      move_init = draw_step
      Up_Enabled = True '拖动画直线
      Move_Enabled = True
  End If
 ElseIf draw_step = 1 Then
      mouse_up_no_enabled = True
ElseIf draw_step = 2 Then
         Call set_select_point
         Call set_select_line
         Call set_select_circle
         Call set_forbid_point
         Call set_forbid_line(temp_line(0))
         Call set_forbid_circle
     If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 0) > 0 Then
'        If operator = "paral_verti" Then
        MDIForm1.StatusBar1.Panels(1).text = LoadResString_(525, _
                          "//1//" & m_poi(temp_point(2).no).data(0).data0.name & _
                          "//2//" & m_poi(temp_point(1).no).data(0).data0.name & _
                                    m_poi(temp_point(2).no).data(0).data0.name) '过作的直线的平行（垂直）线
     End If
            Call init_draw_data
        operator = "draw_point_and_line"
        list_type_for_draw = 1
End If
End If
 '**********************************************
Case "epolygon" 'Then ' draw_step = 3
If list_type_for_draw > 0 Then
 If draw_step = 0 Then
  If draw_new_point(mouse_down_coord, ele1, ele2, red, True, 0) > 0 Then     '
      move_init = draw_step
      Up_Enabled = True '拖动画直线
      Move_Enabled = True
  End If
 ElseIf draw_step = 1 Then
      mouse_up_no_enabled = True
ElseIf draw_step = 2 Then
 If area_triangle(m_poi(temp_point(0).no).data(0).data0.coordinate, _
     m_poi(temp_point(1).no).data(0).data0.coordinate, init_p) > 0 Then
 ty = True
 Else
 ty = False
 End If
  poly_no% = 0
 If list_type_for_draw = 1 Then
   Call set_polygon(temp_point(0).no, temp_point(1).no, 3, _
        poly_no%, ty, True, 0)
 ElseIf list_type_for_draw = 2 Then
   Call set_polygon(temp_point(0).no, temp_point(1).no, 4, _
       poly_no%, ty, True, C_display_wenti.m_last_input_wenti_no)
ElseIf list_type_for_draw = 3 Then
 Call set_polygon(temp_point(0).no, temp_point(1).no, 5, _
     poly_no%, ty, True, 0)
ElseIf list_type_for_draw = 4 Then
Call set_polygon(temp_point(0).no, temp_point(1).no, 6, _
       poly_no%, ty, True, 0)
End If
Call init_draw_data 'If change_operat_statue = True Then
operat_is_acting = False
End If
End If
Case "change_picture" ' Then
If list_type_for_draw = 1 Then
 If change_fig_type = 0 Then
  If draw_step = 0 Then
   operat_is_acting = True
    If set_change_fig > 0 Then
     Exit Sub
    End If
    If Polygon_for_change.p(0).total_v > 0 Then
     draw_step = 2
    Else
'    mdiform1.StatusBar1.Panels(1).text = "设置变换图形"
       temp_point(0).no = read_point(mouse_down_coord, 0)
       'test Call choce_polygon_for_change(temp_point(0))
         draw_step = 1
    End If
  ElseIf draw_step = 1 Then
   temp_point(0).no = read_point(mouse_down_coord, 0)
    If temp_point(0).no > 0 Then
     'test If choce_polygon_for_change(temp_point(0)) Or _
     'test  Polygon_for_change.p(0).total_v = 2 Then
     'test  draw_step = 2
     'test End If
    End If
  ElseIf draw_step = 2 Then
   temp_point(0).no = read_point(mouse_down_coord, 0)
    If temp_point(0).no > 0 And Polygon_for_change.p(0).total_v = 2 Then
      If temp_point(0).no <> Polygon_for_change.p(0).v(0) Then
      'test Call choce_polygon_for_change(temp_point(0))
        draw_step = 1 '直线
      Exit Sub
    Else
     set_change_fig = polygon_
    Exit Sub
   End If
  End If
  End If
' Else
 End If
 ElseIf list_type_for_draw = 2 Then
'test If choce_circle_for_change(read_point(mouse_down_coord, 0), temp_circle(0)) Then
'test  set_change_fig = circle_
'test End If
 ElseIf list_type_for_draw = 3 Or _
      list_type_for_draw = 4 Or list_type_for_draw = 5 Then
  draw_step = 3
 If last_conditions.last_cond(1).change_picture_type = line_ Then
  ' Call draw_change_line(0)
 ElseIf last_conditions.last_cond(1).change_picture_type = polygon_ Then
  'test Call draw_change_polygon(0)
 ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
  'test Call draw_change_circle(0)
'  Draw_form.Line (line_for_move.coord(0).X, _
       line_for_move.coord(0).Y)-(move_x&, move_y&), QBColor(7)
  ' move_x& = X
   ' move_y& = Y
 End If
   'Call draw_change_polygon(0)
  'Call center_display(center_p)
ElseIf list_type_for_draw = 6 Then '画角平分线
 If change_fig_type = 0 Then
  If draw_step = 0 Then
   operat_is_acting = True
    If set_change_fig > 0 Then
     Exit Sub
    End If
' mdiform1.StatusBar1.Panels(1).text = "设置变换图形"
   temp_point(0).no = read_point(mouse_down_coord, 0)
    If temp_point(0).no > 0 Then
     Call C_display_picture.draw_red_point(temp_point(0).no)
     draw_step = 1
    End If
 ElseIf draw_step = 1 Then
  temp_point(1).no = read_point(mouse_down_coord, 0)
   If temp_point(1).no = temp_point(0).no Or temp_point(1).no = 0 Then
    Exit Sub
   Else
    Call C_display_picture.draw_red_point(temp_point(1).no)
   'test If choce_line_for_change(temp_point(0), temp_point(1)) Then
    'test   set_change_fig = line_
    'test    Exit Sub
   'test End If
   End If
 End If
End If
ElseIf list_type_for_draw = 14 Then
  If draw_step = 0 Then
  symmetry_line.p(0).X = mouse_down_coord.X
  symmetry_line.p(0).Y = mouse_down_coord.Y
  symmetry_line.p(1).X = mouse_down_coord.X
  symmetry_line.p(1).Y = mouse_down_coord.Y
  draw_step = 1
  ElseIf draw_step = 1 Then
  symmetry_line.p(1).X = X
  symmetry_line.p(1).Y = Y
  End If
ElseIf list_type_for_draw = 7 Then
 If draw_step = 0 Then
  symmetry_line.p(0).X = X
  symmetry_line.p(0).Y = Y
  symmetry_line.p(1).X = X
  symmetry_line.p(1).Y = Y
  draw_step = 1
  Call center_display(symmetry_line.p(0))
 ElseIf draw_step = 1 Then
  symmetry_line.p(1).X = X
  symmetry_line.p(1).Y = Y
  Call center_display(symmetry_line.p(1))
       Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
         (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
   If Abs(X - symmetry_line.p(0).X) < 5 And Abs(Y - symmetry_line.p(0).Y) < 5 Then
     Exit Sub
   Else
    If set_change_fig = polygon_ Then
    'test Call oxis_symmetry_polygon( _
      Polygon_for_change.p(0), symmetry_line.p(0), symmetry_line.p(1))
     'test  Call draw_change_polygon(1)
     'test      Call link_v_for_change_polygon
     Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
          (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
     Call center_display(symmetry_line.p(0))
     Call center_display(symmetry_line.p(1))
  ElseIf set_change_fig = circle_ Then
    'test Call oxis_symmetry_circle(symmetry_line.p(0), symmetry_line.p(1))
    'test   Call draw_change_circle(1)
    'test       Call link_v_for_change_circle
     Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
          (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
   Call center_display(symmetry_line.p(0))
   Call center_display(symmetry_line.p(1))
 ElseIf set_change_fig = line_ Then
   'test    Call draw_change_line(5)
   'test  Call oxis_symmetry_line(symmetry_line.p(0), symmetry_line.p(1))
   'test    Call draw_change_line(5)
  'test         Call link_v_for_change_line
  End If
  draw_step = 3
 End If
 ElseIf draw_step = 2 Then
  If Abs(X - symmetry_line.p(0).X) < 5 And Abs(Y - symmetry_line.p(0).Y) < 5 Then
  t_coord.X = symmetry_line.p(0).X
  symmetry_line.p(0).X = symmetry_line.p(1).X
  symmetry_line.p(1).X = t_coord.X
  t_coord.Y = symmetry_line.p(0).Y
  symmetry_line.p(0).Y = symmetry_line.p(1).Y
  symmetry_line.p(1).Y = t_coord.Y
  If last_conditions.last_cond(1).change_picture_type = polygon_ Then
  ' Call draw_change_polygon(0)
    Call center_display(symmetry_line.p(0))
     Call center_display(symmetry_line.p(1))
     'test Call link_v_for_change_polygon
       Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
         (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
  ElseIf last_conditions.last_cond(1).change_picture_type = line_ Then
  ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
  'test Call draw_change_circle(0)
    Call center_display(symmetry_line.p(0))
     Call center_display(symmetry_line.p(1))
   'test   Call link_v_for_change_circle
       Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
         (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
  End If
   draw_step = 3
  ElseIf Abs(X - symmetry_line.p(1).X) < 5 And Abs(Y - symmetry_line.p(1).Y) < 5 Then
   If last_conditions.last_cond(1).change_picture_type = polygon_ Then
   'test Call draw_change_polygon(0)
     Call center_display(symmetry_line.p(0))
      Call center_display(symmetry_line.p(1))
      'test Call link_v_for_change_polygon
        Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
         (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
   ElseIf last_conditions.last_cond(1).change_picture_type = line_ Then
   ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
   'test Call draw_change_circle(0)
     Call center_display(symmetry_line.p(0))
      Call center_display(symmetry_line.p(1))
      'test Call link_v_for_change_circle
        Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
         (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
   End If
     draw_step = 3
   End If
  ElseIf Abs(symmetry_line.p(0).X - symmetry_line.p(1).X) < 5 And _
           Abs(symmetry_line.p(0).Y - symmetry_line.p(1).Y) < 5 Then
          symmetry_line.p(1).X = Int(X)
          symmetry_line.p(1).Y = Int(Y)
   draw_step = 3
  End If
ElseIf list_type_for_draw = 8 Then '中心对称
  If draw_step = 0 Then
  center_p.X = Int(X)
  center_p.Y = Int(Y)
  draw_step = 2
  If last_conditions.last_cond(1).change_picture_type = polygon_ Then
  'test Call center_display(center_p)
  'test Call center_symmetry_polygon(Polygon_for_change.p(0), center_p)
    'test Call draw_change_polygon(1)
     Call center_display(center_p)
     'test Call link_v_for_change_polygon
   ElseIf last_conditions.last_cond(1).change_picture_type = line_ Then
   'test Call draw_change_line(5)
     Call center_display(center_p)
     'test Call center_symmetry_line(center_p)
       'test Call draw_change_line(5)
        'test Call link_v_for_change_line
   ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
     'test Call center_symmetry_circle(center_p)
     'test Call draw_change_circle(1)
        Call center_display(center_p)
          'test Call link_v_for_change_circle
   End If
 ElseIf draw_step = 2 Then
   If Abs(X - center_p.X) < 5 And Abs(Y - center_p.Y) < 5 Then
   If last_conditions.last_cond(1).change_picture_type = polygon_ Then
    'test Call draw_change_polygon(0)
     Call center_display(center_p)
     'test Call link_v_for_change_polygon
   ElseIf last_conditions.last_cond(1).change_picture_type = line_ Then
     'Call center_display(center_p)
      'Call draw_change_line(5)
       ' Call link_v_for_change_line
        '   Call center_symmetry_line(center_p)
    'Call center_display(center_p)
    ' Call draw_change_line(5)
     '  Call link_v_for_change_line
   ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
    'test  Call draw_change_circle(1)
     Call center_display(center_p)
    'test Call link_v_for_change_circle
  End If
    draw_step = 3
   End If
  End If
End If
Case "set_view_point" 'Then
If list_type_for_draw = 0 Then
 tp1% = read_point(mouse_down_coord, 0)
 If tp1% > 0 Then
   For i% = 1 To last_conditions.last_cond(1).last_view_point_no
    If view_point(i%).poi = tp1% Then
     Exit Sub
    End If
   Next i%
  last_conditions.last_cond(1).last_view_point_no = _
       last_conditions.last_cond(1).last_view_point_no + 1
   view_point(last_conditions.last_cond(1).last_view_point_no).poi = tp1%
         Call C_display_picture.draw_red_point(tp1%)
         view_point(last_conditions.last_cond(1).last_view_point_no).old_coordinate = _
           m_poi(tp1%).data(0).data0.coordinate
 End If
End If
Case "move_point" 'Then
If list_type_for_draw = 0 Then
Else
yidian_no = read_point(mouse_down_coord, 0)
If yidian_no > 0 And m_poi(yidian_no).data(0).degree > 0 And yidian_type = 2 Then
 C_curve.set_move_point_no (yidian_no)
 '自由点半自由点
 Exit Sub
End If
If list_type_for_draw = 2 Then
  Draw_form.PopupMenu MDIForm1.judge, 0, X + 5, Y + 5
   Exit Sub
ElseIf yidian_type = 0 Or yidian_type = 5 Then
   mouse_move_coord.X = Int(X)
   mouse_move_coord.Y = Int(Y)
If list_type_for_draw = 3 Then
    Call C_curve.set_curve_poi(0, Int(X), Int(Y), 0, 0, 1, yidian_type) '画直线
ElseIf list_type_for_draw = 4 Then
    Call C_curve.set_curve_poi(0, Int(X), Int(Y), 0, 0, 2, yidian_type) '画圆
ElseIf list_type_for_draw = 5 Then
    Call C_curve.set_curve_poi(0, Int(X), Int(Y), 0, 0, 2, yidian_type) '画点列曲线
End If
 If yidian_type = 0 Then
 yidian_type = 1
 End If
End If
Exit Sub
End If
'ElseIf yidian_type = 17 Or yidian_type = 18 Or yidian_type = 19 Then
'   yidian_stop = True
'If Abs(m_poi(yidian_no).data(0).data0.coordinate.X - Int(X)) < 8 And _
 '  Abs(m_poi(yidian_no).data(0).data0.coordinate.Y - Int(Y)) < 8 Then
 ' mouse_move_coord.X = Int(X)
 ' mouse_move_coord.Y = Int(Y)
'MDIForm1.StatusBar1.Panels(1).text = LoadResString_(135)
' change_pic = 2 '暂停
' If yidian_type = 17 Then
'  yidian_type = 25 '直线上移动点
' ElseIf yidian_type = 18 Then
'  yidian_type = 26 '圆上移动点
' ElseIf yidian_type = 19 Then
'  yidian_type = 27 '曲线移动点
' End If
'Else
' If yidian_type = 17 Then
'  Call C_curve.set_curve_poi(Int(X), Int(Y), 0, 0, 3, 5)
' ElseIf yidian_type = 18 Then
'   Call C_curve.set_curve_poi(Int(X), Int(Y), 0, 0, 2, 5)
' End If
'  mouse_move_coord.X = Int(X)
'  mouse_move_coord.Y = Int(Y)
'Exit Sub
'End If
'End If
Case "measure" 'Then
   temp_point(draw_step).no = read_point(mouse_down_coord, 0)
    If temp_point(draw_step).no = 0 Then
     Exit Sub
    Else
      Call C_display_picture.draw_red_point(temp_point(draw_step).no)
    End If
If list_type_for_draw = 1 Then '长度
  If draw_step = 0 Then
        draw_step = 1
  ElseIf draw_step = 1 Then
    If temp_point(1).no <> temp_point(0).no Then
     If is_set_function_data = 0 Or is_set_function_data > 2 Then
      Call set_line_value_for_measure(temp_point(0).no, temp_point(1).no, 0)
     ElseIf is_set_function_data = 1 Then
      function_data.variant_data.poi(0) = temp_point(0).no
      function_data.variant_data.poi(1) = temp_point(0).no
      Call set_menu_for_set_function_data1
      is_set_function_data = 2
     ElseIf is_set_function_data = 2 Then
      function_data.function_data.poi(0) = temp_point(0).no
      function_data.function_data.poi(1) = temp_point(0).no
      Call recove_set_menu_for_set_function_data
      is_set_function_data = 3
     End If
'      Call C_display_picture.redraw_point(temp_point(0))
'      Call C_display_picture.redraw_point(temp_point(1))
     draw_step = 0
   Else
  '   Call C_display_picture.redraw_point(temp_point(1))
   End If
 End If
ElseIf list_type_for_draw = 2 Then '长度
  If draw_step = 0 Then
        draw_step = 1
   ElseIf draw_step = 1 Then
    If temp_point(1).no <> temp_point(0).no Then
        draw_step = 2
     Else
 '     Call C_display_picture.redraw_point(temp_point(1))
     End If
   ElseIf draw_step = 2 Then
    If temp_point(2).no <> temp_point(0).no And temp_point(2).no <> temp_point(1).no Then
     If is_set_function_data = 0 Or is_set_function_data > 2 Then
      Call set_distence_p_line_for_measure(temp_point(0).no, _
            temp_point(1).no, temp_point(2).no, 0)
     ElseIf is_set_function_data = 1 Then
      temp_line(0) = line_number(temp_point(1).no, temp_point(2).no, _
                                 pointapi0, pointapi0, _
                                 depend_condition(point_, temp_point(1).no), _
                                 depend_condition(point_, temp_point(2).no), _
                                 condition, condition_color, 1, 1)
      function_data.variant_data.poi(0) = temp_point(0).no
      function_data.variant_data.poi(1) = temp_line(0) + 1000
      Call set_menu_for_set_function_data1
      is_set_function_data = 2
     ElseIf is_set_function_data = 2 Then
      temp_line(0) = line_number(temp_point(1).no, temp_point(2).no, _
                                 pointapi0, pointapi0, _
                                 depend_condition(point_, temp_point(1).no), _
                                 depend_condition(point_, temp_point(2).no), _
                                 condition, condition_color, 1, 1)
      function_data.function_data.poi(0) = temp_point(0).no
      function_data.function_data.poi(1) = temp_line(0) + 1000
      Call recove_set_menu_for_set_function_data
      is_set_function_data = 3
     End If
'          Call C_display_picture.redraw_point(temp_point(0))
'     Call C_display_picture.redraw_point(temp_point(1))
'     Call C_display_picture.redraw_point(temp_point(2))
         draw_step = 0
  Else
'        Call C_display_picture.redraw_point(temp_point(2))
  End If
  End If
 ElseIf list_type_for_draw = 3 Then '长度
   If draw_step = 0 Then
    temp_polygon.v(0) = temp_point(0).no
     temp_polygon.total_v = 1
'        Call C_display_picture.draw_red_point(temp_point(0))
        draw_step = 1
   ElseIf draw_step > 0 Then
     If temp_point(1).no <> temp_polygon.v(0) Then
       temp_polygon.v(temp_polygon.total_v) = temp_point(1).no
'         Call C_display_picture.draw_red_point(temp_polygon.v(temp_polygon.total_v))
        Call draw_red_line(line_number0( _
           temp_polygon.v(temp_polygon.total_v), _
             temp_polygon.v(temp_polygon.total_v - 1), 0, 0))
              temp_polygon.total_v = temp_polygon.total_v + 1
     Else
         Call draw_red_line(line_number0( _
           temp_polygon.v(temp_polygon.total_v - 1), _
              temp_polygon.v(0), 0, 0))
For i% = 1 To temp_polygon.total_v - 1
  Call redraw_red_line(line_number0(temp_polygon.v(i%), _
              temp_polygon.v(i% - 1), 0, 0))
'   Call C_display_picture.redraw_point(temp_polygon.v(i%))
Next i%
  Call redraw_red_line(line_number0(temp_polygon.v(0), _
         temp_polygon.v(temp_polygon.total_v - 1), 0, 0))
 '  Call C_display_picture.redraw_point(temp_polygon.v(0))
  If is_set_function_data = 0 Or is_set_function_data = 3 Then
  Call set_area_of_polygon_for_measure(temp_polygon, 0)
  ElseIf is_set_function_data = 1 Then
      Call set_menu_for_set_function_data1
      is_set_function_data = 2
  ElseIf is_set_function_data = 2 Then
      Call recove_set_menu_for_set_function_data
      is_set_function_data = 3
  End If
        Call init_draw_data
   End If
  End If
  ElseIf list_type_for_draw = 4 Then ' 角度
  If draw_step = 0 Then
        draw_step = 1
   ElseIf draw_step = 1 Then
    If temp_point(0).no <> temp_point(1).no Then
'       Call C_display_picture.draw_red_point(temp_point(1))
        draw_step = 2
    Else
 '     Call C_display_picture.redraw_point(temp_point(1))
    End If
   ElseIf draw_step = 2 Then
    If temp_point(2).no <> temp_point(1).no And temp_point(2).no <> temp_point(0).no Then
     If is_new_angle_value_for_measur(temp_point(0).no, temp_point(1).no, _
          temp_point(2).no, k%) Then
      Call value_of_angle( _
          angle_value_for_measur(k%))
    last_measur_string = last_measur_string + 1
    Measur_string(last_measur_string) = set_display_angle0( _
     m_poi(angle_value_for_measur(k%).poi(0)).data(0).data0.name + _
      m_poi(angle_value_for_measur(k%).poi(1)).data(0).data0.name + _
       m_poi(angle_value_for_measur(k%).poi(2)).data(0).data0.name) + _
       "=" + angle_value_for_measur(k%).value
'     Call C_display_picture.redraw_point(temp_point(0))
'     Call C_display_picture.redraw_point(temp_point(1))
'     Call C_display_picture.redraw_point(temp_point(2))
    If is_set_function_data = 0 Or is_set_function_data = 3 Then
    Call display_m_string(last_measur_string, display)
     angle_value_for_measur(k%).string_no = last_measur_string
    ElseIf is_set_function_data = 1 Then
    ElseIf is_set_function_data = 2 Then
    End If
        draw_step = 0
    Else
 '    Call C_display_picture.redraw_point(temp_point(0))
 '    Call C_display_picture.redraw_point(temp_point(1))
 '    Call C_display_picture.redraw_point(temp_point(2))
    Call display_m_string(last_measur_string, display)
        draw_step = 0
   End If
  Else
' Call C_display_picture.redraw_point(temp_point(2))
End If
End If
End If
Case "set" 'Then
   temp_point(draw_step).no = read_point(mouse_down_coord, 0)
    If temp_point(draw_step).no = 0 Then
     Exit Sub
    Else
      Call C_display_picture.draw_red_point(temp_point(draw_step).no)
    End If
If list_type_for_draw = 2 Then
  If draw_step = 0 Then
        draw_step = 1
  ElseIf draw_step = 1 Then
    If temp_point(1).no <> temp_point(0).no Then
     Call set_line_value_for_measure(temp_point(0).no, temp_point(1).no, _
          set_measure_no%)
'      Call C_display_picture.redraw_point(temp_point(0))
'      Call C_display_picture.redraw_point(temp_point(1))
     draw_step = 2
   Else
'     Call C_display_picture.redraw_point(temp_point(1))
   End If
 End If
 ElseIf list_type_for_draw = 3 Then '长度
    If draw_step = 0 Then
        draw_step = 1
   ElseIf draw_step = 1 Then
    If temp_point(1).no <> temp_point(0).no Then
        draw_step = 2
     Else
'      Call C_display_picture.redraw_point(temp_point(1))
     End If
   ElseIf draw_step = 2 Then
    If temp_point(2).no <> temp_point(0).no And temp_point(2).no <> temp_point(1).no Then
     Call set_distence_p_line_for_measure(temp_point(0).no, _
            temp_point(1).no, temp_point(2).no, set_measure_no%)
'     Call C_display_picture.redraw_point(temp_point(0))
'     Call C_display_picture.redraw_point(temp_point(1))
'     Call C_display_picture.redraw_point(temp_point(2))
         draw_step = 3
  Else
'        Call C_display_picture.redraw_point(temp_point(2))

  End If
  End If
ElseIf list_type_for_draw = 4 Then '长度
  If draw_step = 0 Then
    temp_polygon.v(0) = temp_point(0).no
     temp_polygon.total_v = 1
        Call C_display_picture.draw_red_point(temp_point(0).no)
        draw_step = 1
   ElseIf draw_step > 0 Then
     If temp_point(1).no <> temp_polygon.v(0) Then
       temp_polygon.v(temp_polygon.total_v) = temp_point(1).no
         Call C_display_picture.draw_red_point(temp_polygon.v(temp_polygon.total_v))
        Call draw_red_line(line_number0( _
           temp_polygon.v(temp_polygon.total_v), _
             temp_polygon.v(temp_polygon.total_v - 1), 0, 0))
            temp_polygon.total_v = temp_polygon.total_v + 1
     Else
         Call draw_red_line(line_number0( _
           temp_polygon.v(temp_polygon.total_v - 1), _
              temp_polygon.v(0), 0, 0))
For i% = 1 To temp_polygon.total_v - 1
  Call redraw_red_line(line_number0(temp_polygon.v(i%), _
              temp_polygon.v(i% - 1), 0, 0))
'   Call C_display_picture.redraw_point(temp_polygon.v(i%))
Next i%
  Call redraw_red_line(line_number0(temp_polygon.v(0), _
         temp_polygon.v(temp_polygon.total_v - 1), 0, 0))
'   Call C_display_picture.redraw_point(temp_polygon.v(0))
  Call set_area_of_polygon_for_measure(temp_polygon, set_measure_no%)
        Call init_draw_data
   End If
  End If
End If
Case "re_name"
If list_type_for_draw = 3 Then
 'Call remove_uncomplete_operat(old_operator)
  temp_point(0).no = read_point(mouse_down_coord, 0)
If temp_point(0).no > 0 Then
 Draw_form.PopupMenu MDIForm1.judge, 0, X + 5, Y + 5
'   "取消该点" "保留该点"
     Call C_display_picture.draw_red_point(temp_point(draw_step).no)   ', display)
End If
Else
 temp_point(0).no = read_point(mouse_down_coord, 0) ' 读点
If list_type_for_draw = 1 Then
  If temp_point(0).no = 0 Or m_poi(temp_point(0).no).data(0).data0.name <> "" Then
   Exit Sub
  ElseIf choose_point > 0 Then '以选点
   yidian_stop = True
'    Call C_display_picture.redraw_point(choose_point)
  End If
Else
 If temp_point(0).no = 0 Then
  Exit Sub
 End If
End If
   choose_point = temp_point(0).no
    If choose_point > 0 Then '新点
      Call C_display_picture.draw_red_point(choose_point)
      yidian_stop = False
       Call C_display_picture.flash_point(choose_point)
    End If
End If
End Select
End If
End If

End Sub


Public Sub plane_geometry_draw_mouse_move(Button As Integer, Shift As Integer, X As Single, Y As Single)
Draw_form.Label2.visible = False
Dim f As Integer
Dim c As Single
Dim l As Single
Dim k1 As Single
Dim k2 As Single
Dim A As Single
Dim i%, si!, co!, r!, dis&
Dim t_s!, t_c!
Dim ch As String
Dim n As Byte
Dim p As POINTAPI
Dim p1 As POINTAPI
Dim p2 As POINTAPI
Dim p3 As POINTAPI
Dim p4 As POINTAPI
Dim p5 As POINTAPI
Dim p6 As POINTAPI
Dim p7 As POINTAPI
Dim p8 As POINTAPI
Dim e(2) As Integer
Dim ele(1) As condition_type
   If (Abs(mouse_down_coord.X - X) < 5 And Abs(mouse_down_coord.Y - Y) < 5) Then
      Exit Sub '未移动
   End If
'********************************************************************
mouse_move_coord.X = X
mouse_move_coord.Y = Y '移动坐标
p.X = Int(X)
p.Y = Int(Y)
'******************************************************************************************
'*****************************************************************************************
If Button = 2 Then '右键，移动全图
   If operator = "draw_point_and_line" Or operator = "draw_circle" Or operator = "paral_verti" Or _
        operator = "verti_midpoint" Then
    If is_same_POINTAPI(mouse_move_coord, mouse_down_coord) = False Then '
     Call change_picture_from_move(0, minus_POINTAPI(p, mouse_move_coord))
    End If
   End If
'*************************************************************************************
'****************************************************************************************
ElseIf Button = 0 Then '不按键
  If operator = "draw_circle" Then
     If (list_type_for_draw = 1 Or list_type_for_draw = 2) Then
      If m_temp_circle_for_input.is_using And draw_step = 0 Then
     ' m_temp_circle_for_input.data(0).center = mouse_down_coord
       m_temp_circle_for_input.data(0).radii = distance_of_two_POINTAPI( _
        m_temp_circle_for_input.data(0).c_coord, mouse_move_coord)
'         m_temp_circle_for_input.is_using = True
           list_type_for_draw = 2
            Call draw_temp_circle_for_input
     ElseIf draw_step = 1 Then
      m_temp_circle_for_input.data(0).radii = _
        circle_radii0(m_poi(temp_point(0).no).data(0).data0.coordinate, m_poi(temp_point(1).no).data(0).data0.coordinate, _
           mouse_move_coord, m_temp_circle_for_input.data(0).c_coord)
      Call draw_temp_circle_for_input
     End If
    ElseIf list_type_for_draw = 5 And draw_step = 2 Then '画与已知相切的圆
     If temp_circle(0) > 0 Then
     dis& = distance_of_two_POINTAPI(mouse_move_coord, m_Circ(temp_circle(0)).data(0).data0.c_coord)
      Call inter_point_line_circle3(mouse_move_coord, paral_, mouse_move_coord, m_Circ(temp_circle(0)).data(0).data0.c_coord, _
              m_Circ(temp_circle(0)).data(0).data0, p1, 0, p2, 0, 0, 0)
         temp_tangent_circle_no(0) = set_tangent_circle_data(mouse_move_coord, dis& + m_Circ(temp_circle(0)).data(0).data0.radii, _
               0, p1, depend_condition(circle_, temp_circle(0)), 0, pointapi0, depend_condition(0, 0), new_point_on_line_circle12, _
                  temp_tangent_circle_no(0))
         temp_tangent_circle_no(1) = set_tangent_circle_data(mouse_move_coord, Abs(dis& - m_Circ(temp_circle(0)).data(0).data0.radii), _
               0, p2, depend_condition(circle_, temp_circle(0)), 0, pointapi0, depend_condition(0, 0), new_point_on_line_circle21, _
                  temp_tangent_circle_no(1))
               
     ElseIf temp_lines_for_draw(0) = 1 Then
      If distance_point_to_line(mouse_move_coord, _
              m_poi(m_lin(temp_lines_for_draw(1)).data(0).data0.poi(0)).data(0).data0.coordinate, paral_, _
               m_poi(m_lin(temp_lines_for_draw(1)).data(0).data0.poi(0)).data(0).data0.coordinate, _
                m_poi(m_lin(temp_lines_for_draw(1)).data(0).data0.poi(1)).data(0).data0.coordinate, _
                 dis&, p) Then
               temp_tangent_circle_no(0) = set_tangent_circle_data(mouse_move_coord, Abs(dis&), _
                 0, p, depend_condition(line_, temp_lines_for_draw(1)), 0, pointapi0, depend_condition(0, 0), 0)
      End If
     End If
    ElseIf list_type_for_draw = 6 And draw_step = 5 Then
     dis& = distance_of_two_POINTAPI(mouse_move_coord, m_Circ(temp_circle(0)).data(0).data0.c_coord) - _
              m_Circ(temp_circle(0)).data(0).data0.radii
     p5 = inter_point_circle_circle_by_pointapi(m_Circ(temp_circle(0)).data(0).data0.c_coord, _
            m_Circ(temp_circle(0)).data(0).data0.radii + dis&, _
             m_Circ(temp_circle(1)).data(0).data0.c_coord, _
            m_Circ(temp_circle(1)).data(0).data0.radii + dis&, p1, p2)
     p6 = inter_point_circle_circle_by_pointapi(m_Circ(temp_circle(0)).data(0).data0.c_coord, _
            m_Circ(temp_circle(0)).data(0).data0.radii + dis&, _
             m_Circ(temp_circle(1)).data(0).data0.c_coord, _
             m_Circ(temp_circle(1)).data(0).data0.radii - dis&, p3, p4)
     If p5.X <> -10000 Then
        If distance_of_two_POINTAPI(mouse_move_coord, p1) <= _
            distance_of_two_POINTAPI(mouse_move_coord, p2) Then
             p7 = p1
        Else
             p7 = p2
        End If
     End If
     If p6.X <> -10000 Then
        If distance_of_two_POINTAPI(mouse_move_coord, p3) <= _
              distance_of_two_POINTAPI(mouse_move_coord, p4) Then
             p8 = p3
        Else
             p8 = p4
        End If
     End If
     If p5.X <> -10000 And p6.X <> -10000 Then
         If distance_of_two_POINTAPI(mouse_move_coord, p7) <= _
            distance_of_two_POINTAPI(mouse_move_coord, p8) Then
             p7 = p7
         Else
              p7 = p8
         End If
     ElseIf p5.X <> -10000 Then
     ElseIf p6.X <> -10000 Then
            p7 = p8
     Else
         dis& = 0
     End If
        temp_tangent_circle_no(0) = set_tangent_circle_data(p7, Abs(dis&), _
                  0, pointapi0, depend_condition(circle_, temp_circle(0)), 0, pointapi0, _
                   depend_condition(circle_, temp_circle(1)), 0, 0)
    End If
 End If
 If display_temp_four_point_fig = 1 Then
  If temp_four_point_fig_type = long_squre_ Then
   Call draw_temp_long_squre(p, 0)
   Call draw_temp_long_squre(mouse_move_coord, 0)
  ElseIf temp_four_point_fig_type = parallelogram_ Then
   Call draw_temp_parallelogram(p, 0)
   Call draw_temp_parallelogram(mouse_move_coord, 0)
  ElseIf temp_four_point_fig_type = equal_side_tixing_ Then
   Call draw_temp_equal_side_tixing(p, 0)
   Call draw_temp_equal_side_tixing(mouse_move_coord, 0)
  ElseIf temp_four_point_fig_type = rhombus_ Then
   Call draw_temp_rhombus(p, 0)
   Call draw_temp_rhombus(mouse_move_coord, 0)
  End If
   'mouse_move_coord = p
 End If
'**************************************************************************************************
'mouse_move_coord = p
Call read_inter_point(mouse_move_coord, ele(0), ele(1), 0, False, 0) '画辅助线
'**************************************************************************************
If operator = "draw_circle" Then
 If list_type_for_draw = 2 Then
  If Move_Enabled Then  '无键移动
    If Move_statue <> 2 Then
     MDIForm1.StatusBar1.Panels(1).text = _
      LoadResString_(2025, "\\1\\" + m_poi(temp_point(0).no).data(0).data0.name)
      Move_statue = 2
 '     draw_step = draw_step + 1 '画线
    End If
  End If
 End If
End If
'********************************************************************************************
'***********************************************************************************************
ElseIf Button = 1 Then
event_statue = mouse_is_moving
If input_text_statue Or _
    event_statue = wait_for_modify_char Or _
     event_statue = wait_for_input_char Or _
        event_statue = input_char_again Or _
         draw_statue = False Then

 '文本输入状态
 Exit Sub
End If
'Call draw_aid_line(p, 1)
If operator = "draw_point_and_line" Or operator = "paral_verti" Or operator = "verti_midpoint" Or operator = "epolygon" Then
 If Move_Enabled Then  '按左键移动
    If m_temp_line_for_input.is_using = False Then '未进入拖动划线,如已进入,跳过这段程序
            m_temp_line_for_input.is_using = True '
            m_temp_line_for_input.data(0).total_color = conclusion_color
            m_temp_line_for_input.data(0).poi(0) = temp_point(move_init).no
            m_temp_line_for_input.data(0).end_point_coord(0) = mouse_down_coord
            m_temp_line_for_input.data(0).end_point_coord(1) = mouse_move_coord
            m_temp_line_for_input.data(0).visible = 1
            Call draw_temp_line_for_input
    Else
           m_temp_line_for_input.data(0).end_point_coord(1) = mouse_move_coord
           Call draw_temp_line_for_input
    ' Call C_display_picture.draw_line(temp_line(draw_line_no), mouse_move_coord.X, mouse_move_coord.Y)
    End If
    Call read_inter_point(mouse_move_coord, ele1, ele2, 0, False, 0)
  'End If
 End If
ElseIf operator = "draw_circle" Then '画圆
If list_type_for_draw = 0 Then
Exit Sub
ElseIf list_type_for_draw = 1 Then
If Move_Enabled And draw_step = 0 Then
 m_temp_circle_for_input.data(0).radii = distance_of_two_POINTAPI( _
         m_temp_circle_for_input.data(0).c_coord, mouse_move_coord)
         Call draw_temp_circle_for_input
         Move_statue = 1
End If
ElseIf list_type_for_draw = 2 Then '画三点圆的第三点
  If draw_step = 2 Then
   If Move_Enabled Then
     If is_same_POINTAPI(m_poi(temp_point(move_init).no).data(0).data0.coordinate, mouse_move_coord) = False Then
        If Move_statue <> 1 Then
         Move_statue = 1
         ' draw_step = draw_step + 1 '画线
        End If
     Call C_display_picture.draw_circle(temp_circle(0), mouse_move_coord.X, mouse_move_coord.Y)
 End If
End If
  End If
ElseIf list_type_for_draw = 3 Then
 If draw_step = 5 Then
    input_point_type% = read_inter_point(mouse_move_coord, ele1, ele2, -1000, True, 1)      '选指定点
     Call draw_temp_line_for_move(input_point_type%, ele1.no, ele2.no)
     If tangent_circle_type = 1 Then
     Call set_tangent_circles(m_Circ(temp_circle(0)).data(0).data0, _
                m_poi(m_Circ(temp_circle(0)).data(0).data0.in_point(2)).data(0).data0.coordinate, p1, _
                 m_Circ(temp_circle(0)).data(0).data0.radii, 0, m_Circ(temp_circle(0)).data(0).data0.c_coord, pointapi0, _
                   t_coord1, m_Circ(temp_circle(0)).data(0).data0.in_point(1), t_coord2, 0, 0, False)
     ElseIf tangent_circle_type = 2 Then
     Call set_tangent_circles(m_Circ(temp_circle(0)).data(0).data0, _
                m_poi(m_Circ(temp_circle(0)).data(0).data0.in_point(2)).data(0).data0.coordinate, p1, _
                  0, m_Circ(temp_circle(0)).data(0).data0.radii, pointapi0, m_Circ(temp_circle(0)).data(0).data0.c_coord, _
                   t_coord1, m_Circ(temp_circle(0)).data(0).data0.in_point(1), t_coord2, 0, 0, False)
     End If
    End If
Else
End If
ElseIf operator = "move_point" Then
If list_type_for_draw = 0 Then
 Exit Sub
End If
  If (list_type_for_draw = 3 Or list_type_for_draw = 4 Or list_type_for_draw = 5) And yidian_type = 1 Then
    Call C_curve.set_curve_poi(0, Int(X), Int(Y), 0, 0, 0, yidian_type) '画直线
    Exit Sub
  End If
    If yidian_no > 0 Then
     If m_poi(yidian_no).data(0).parent.co_degree < 2 Then
      If yidian_statue = False Then
       yidian_statue = True
        m_poi(yidian_no%).data(0).data0.coordinate = mouse_move_coord
         'm_poi(yidian_no%).data(0).is_change = True
        Call change_m_point(yidian_no, True)
'        Call change_picture_from_move(yidian_no, mouse_move_coord)
'    Call change_picture_(yidian_no, 0)
'     Call draw_again0(Draw_form, 1)
        yidian_statue = False
     End If
    End If
  End If
'*******
ElseIf operator = "change_picture" Then
If draw_step = 1 Then
 If list_type_for_draw = 7 Then
  Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
        (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
  Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
        (Int(X), Int(Y)), QBColor(12)
   symmetry_line.p(1).X = Int(X)
    symmetry_line.p(1).Y = Int(Y)
 End If
ElseIf draw_step = 3 Then
If change_fig_type = 0 Then
If list_type_for_draw = 3 Then '平移
 If last_conditions.last_cond(1).change_picture_type = line_ Then
   'test Call change_line(p.X - init_p.X, p.Y - init_p.Y, 0, line_for_change.similar_ratio)
  init_p = p
 ElseIf last_conditions.last_cond(1).change_picture_type = polygon_ Then
  'test Call change_polygon(minus_pointapi(p, init_p), 0, _
                                 Polygon_for_change.similar_ratio)
  init_p = p
 ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
 'test Call change_circle(p.X - init_p.X, p.Y - init_p.Y, 0, Circle_for_change.similar_ratio)
  init_p = p
 End If
ElseIf list_type_for_draw = 4 Then '旋转
 If last_conditions.last_cond(1).change_picture_type = line_ Then
     If (init_p.X - line_for_change.line_no(0).center(1).X) * _
        (Int(Y) - line_for_change.line_no(0).center(1).Y) - _
     (init_p.Y - line_for_change.line_no(0).center(1).Y) * _
       (Int(X) - line_for_change.line_no(0).center(1).X) > 0 Then
   'test  Call change_line(0, 0, 0.02, line_for_change.similar_ratio)
       Else
    'test  Call change_line(0, 0, -0.02, line_for_change.similar_ratio)
     End If
     init_p = p
 ElseIf last_conditions.last_cond(1).change_picture_type = polygon_ Then
  If (init_p.X - Polygon_for_change.p(0).coord_center.X) * _
        (Int(Y) - Polygon_for_change.p(0).coord_center.Y) - _
     (init_p.Y - Polygon_for_change.p(0).coord_center.Y) * _
       (Int(X) - Polygon_for_change.p(0).coord_center.X) > 0 Then
   '    yidian_stop = False
    t_coord.X = 0
    t_coord.Y = 0
   'test  Call change_polygon(t_coord, 0.02, Polygon_for_change.similar_ratio)
  Else
    t_coord.X = 0
    t_coord.Y = 0
   'test Call change_polygon(t_coord, -0.02, Polygon_for_change.similar_ratio)
  End If
init_p = p
ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
End If
ElseIf list_type_for_draw = 5 Then '放大
 If last_conditions.last_cond(1).change_picture_type = line_ Then
  If (Int(X) - init_p.X) * (line_for_change.line_no(0).center(1).X - init_p.X) + _
       (Int(Y) - init_p.Y) * (line_for_change.line_no(0).center(1).Y - init_p.Y) > 0 Then
   If line_for_change.similar_ratio > 0.04 Then
    line_for_change.similar_ratio = _
     line_for_change.similar_ratio - 0.02
   End If
   'test Call change_line(0, 0, 0, line_for_change.similar_ratio)
  Else
    line_for_change.similar_ratio = _
     line_for_change.similar_ratio + 0.02
   'test Call change_line(0, 0, 0, line_for_change.similar_ratio)
  End If
 ElseIf last_conditions.last_cond(1).change_picture_type = polygon_ Then
  If (Int(X) - init_p.X) * (Polygon_for_change.p(0).coord_center.X - init_p.X) + _
       (Int(Y) - init_p.Y) * (Polygon_for_change.p(0).coord_center.Y - init_p.Y) > 0 Then
   If Polygon_for_change.similar_ratio > 0.04 Then
    Polygon_for_change.similar_ratio = _
     Polygon_for_change.similar_ratio - 0.02
   End If
    t_coord.X = 0
    t_coord.Y = 0
  'test  Call change_polygon(t_coord, 0, Polygon_for_change.similar_ratio)
  Else
    Polygon_for_change.similar_ratio = _
     Polygon_for_change.similar_ratio + 0.02
    t_coord.X = 0
    t_coord.Y = 0
  'test  Call change_polygon(t_coord, 0, Polygon_for_change.similar_ratio)
  End If
 ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
     If (Int(X) - init_p.X) * (Circle_for_change.c_coord.X - init_p.X) + _
       (Int(Y) - init_p.Y) * (Circle_for_change.c_coord.Y - init_p.Y) > 0 Then
    If Circle_for_change.similar_ratio > 0.04 Then
     Circle_for_change.similar_ratio = _
      Circle_for_change.similar_ratio - 0.02
    End If
 'test  Call change_circle(0, 0, 0, Circle_for_change.similar_ratio)
  Else
    Circle_for_change.similar_ratio = _
     Circle_for_change.similar_ratio + 0.02
  'test Call change_circle(0, 0, 0, Circle_for_change.similar_ratio)
  End If
 End If
  'test   Call draw_change_circle(0)
'Call change_circle(X - move_x&, Y - move_y)
'  Draw_form.Line (line_for_move.coord(0).X, _
       line_for_move.coord(0).Y)-(move_x&, move_y&), QBColor(7)
 '  move_x& = X
  '  move_y& = Y
'  Draw_form.Line (line_for_move.coord(0).X, _
       line_for_move.coord(0).Y)-(move_x&, move_y&), QBColor(7)
ElseIf list_type_for_draw = 7 Then '轴对称
  Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
        (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
   Call center_display(symmetry_line.p(1))
    symmetry_line.p(1).X = Int(X)
    symmetry_line.p(1).Y = Int(Y)
  If last_conditions.last_cond(1).change_picture_type = polygon_ Then
    Call draw_polygon(Polygon_for_change.p(0), Polygon_for_change.direction)
   'test  Call link_v_for_change_polygon
    'test Call oxis_symmetry_polygon(Polygon_for_change.p(0), _
        symmetry_line.p(0), symmetry_line.p(1))
    Call draw_polygon(Polygon_for_change.p(0), Polygon_for_change.direction)
  'test  Call link_v_for_change_polygon
    Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
        (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
    Call center_display(symmetry_line.p(1))
  ElseIf last_conditions.last_cond(1).change_picture_type = line_ Then
    'test    Call draw_change_line(5)
    'test    Call link_v_for_change_line
    'test Call oxis_symmetry_line(symmetry_line.p(0), symmetry_line.p(1))
    'test   Call draw_change_line(5)
    'test    Call link_v_for_change_line
        Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
        (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
          Call center_display(symmetry_line.p(1))
  ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
   'test Call oxis_symmetry_circle(symmetry_line.p(0), symmetry_line.p(1))
   'test Call draw_change_circle(0)
   'test Call link_v_for_change_circle
    Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
        (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
    Call center_display(symmetry_line.p(0))
    Call center_display(symmetry_line.p(1))
  End If
ElseIf list_type_for_draw = 8 Then '中心对称
    Call center_display(center_p)
    center_p.X = Int(X)
    center_p.Y = Int(Y)
    Call center_display(center_p)
  If last_conditions.last_cond(1).change_picture_type = polygon_ Then
   'test Call link_v_for_change_polygon
    Call draw_polygon(Polygon_for_change.p(0), Polygon_for_change.direction)
   'test Call center_symmetry_polygon(Polygon_for_change.p(0), center_p)
   'test Call link_v_for_change_polygon
    Call draw_polygon(Polygon_for_change.p(0), Polygon_for_change.direction)
  ElseIf last_conditions.last_cond(1).change_picture_type = line_ Then
      'test  Call draw_change_line(5)
      'test  Call link_v_for_change_line
    'test Call center_symmetry_line(center_p)
   'test    Call draw_change_line(5)
    'test    Call link_v_for_change_line
  ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
    'test Call center_symmetry_circle(center_p)
     'test Call draw_change_circle(0)
       Call center_display(center_p)
       'test Call link_v_for_change_circle
 End If
End If
ElseIf change_fig_type = 1 Then
End If
End If
'End If
ElseIf operator = "move_paral" Then '平移
'test Call change_polygon(minus_pointapi(p, init_p), 0, 1)
init_p = p
End If
End If

End Sub
Public Sub plane_geometry_draw_mouse_up(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i%, j%, tp1%, tp2%, tp%, tangent_linee%
Dim l%, H%
Dim ch$, A!
Dim ang(1) As Integer
Dim n As Byte
Dim r!
Dim D1&
Dim c_data0 As condition_data_type
Dim di_r(1) As Long
Dim c_coord(1) As POINTAPI
Dim ty As Boolean
Dim re As total_record_type
Dim vf As POINTAPI
If input_text_statue Or _
    event_statue = wait_for_modify_char Or _
     event_statue = wait_for_input_char Or _
      event_statue = input_char_again Or _
       event_statue = wait_for_input_condition Or _
         draw_statue = False Or draw_operate = False Then
'文本输入状态
 Exit Sub
End If
'***************************
'设置输入鼠标的坐标
mouse_up_coord.X = Int(X)
mouse_up_coord.Y = Int(Y)
'*************************************
If Button = 2 Then '按右键
 If operator = "draw_circle" And list_type_for_draw = 1 Then
  If draw_step = 2 Then
    Call init_draw_data
     operat_is_acting = False
  End If
 End If
ElseIf Button = 1 Then '按左键
If event_statue = wait_for_draw_point Then
event_statue = draw_point_up
ElseIf event_statue = mouse_is_moving Then
event_statue = ready
End If
'*******************************************************************************************
'左键抬起后,各种输入菜单的操作
If mouse_up_no_enabled Then
mouse_up_coord = mouse_down_coord
  mouse_up_no_enabled = False
ElseIf (is_same_POINTAPI(mouse_down_coord, mouse_up_coord) And Move_Enabled) Or draw_step Mod 2 = 1 Then
  If (operator = "draw_circle" And list_type_for_draw = 1) Then  '画两点圆，鼠标未移动，转化为画三点圆
         m_Circ(temp_circle(0)).data(0).input_type = aid_condition
         list_type_for_draw = 2
  End If
ElseIf is_same_POINTAPI(mouse_down_coord, mouse_up_coord) Then
   mouse_up_no_enabled = False
Else
   draw_step = draw_step + 1
End If
Select Case operator
Case "draw_point_and_line" '画点线
Select Case list_type_for_draw
 Case 1 '画线或点
       If draw_step = 0 Then
                Call init_draw_data
       ElseIf draw_step = 1 Then
         Call draw_new_point(mouse_up_coord, ele1, ele2, red, True, 255)
         Call init_draw_data
       End If
 Case 2, 3, 4
      If draw_step = 1 Then
         Call draw_temp_line_for_mouse_up(True)
         Call init_draw_data
      End If
 Case 5
    If draw_step = 1 Then
       Call draw_temp_line_for_mouse_up(False)
           MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1975, "")
    ElseIf draw_step = 3 Then
     Call draw_new_point(mouse_up_coord, ele1, ele2, red, True, 255)
     '  Call draw_temp_line_for_mouse_up(True, 0, 0, temp_circle(0))
        m_Circ(temp_circle(0)).data(0).data0.visible = False
         Call C_display_picture.set_m_circle_visible(temp_circle(0), 0)
                Call init_draw_data
    End If
Case 6 '画角平分线
    If draw_step = 1 Then
       Call draw_temp_line_for_mouse_up(False)
    ElseIf draw_step = 3 Then
       Call draw_temp_line_for_mouse_up(False)
          temp_line(draw_line_no) = line_sp_angle(temp_point(0).no, temp_point(1).no, temp_point(3).no)
    End If
End Select
'**************************************************************************************************
Case "draw_circle" 'Then
If list_type_for_draw = 0 Then
 Exit Sub
ElseIf list_type_for_draw = 1 Then
  If draw_step = 1 Then
       '画第二个点,
         Call set_select_point
         Call set_select_line
         Call set_select_circle(temp_circle(0))
         Call set_forbid_point(temp_point(0).no, temp_point(1).no)
         Call set_forbid_line
         Call set_forbid_circle
        m_temp_circle_for_input.is_using = True
        m_temp_circle_for_input.data(0).color = 15 'condition_color
        m_temp_circle_for_input.data(0).center = temp_point(0).no
            Call draw_temp_circle_for_input
     If draw_new_point(mouse_up_coord, ele1, ele2, red, True, 255) > 0 Then
        'temp_circle(0) = m_circle_number(1, 0, m_temp_circle_for_input.data(0).c_coord, temp_point(0).no, _
                        temp_point(1).no, temp_point(2).no, m_temp_circle_for_input.data(0).radii, _
                         0, 0, 0, 1, 1, condition_color, True)
        Call init_draw_data
        operator = "draw_point_and_line"
         list_type_for_draw = 1
     Else
        draw_step = 1
     End If
   End If
 ElseIf list_type_for_draw = 2 Then '画三点圆
 ElseIf list_type_for_draw = 5 Then
   If draw_step = 3 Then
   End If
 End If

'*************************************************************************************
Case "paral_verti" 'Then
   If list_type_for_draw = 1 Then
    If draw_step = 1 Then
             Call draw_temp_line_for_mouse_up(False)
    End If
'*************************************************************************************************
   ElseIf list_type_for_draw = 2 Then
    '画出标准线
        If draw_step = 1 Then
            Call draw_temp_line_for_mouse_up(False)
         temp_point(2).no = m_point_number(mid_POINTAPI( _
                         m_poi(temp_point(0).no).data(0).data0.coordinate, _
                          m_poi(temp_point(1).no).data(0).data0.coordinate), condition, 1, condition_color, "", _
                            depend_condition(point_, temp_point(0).no), depend_condition(point_, temp_point(1).no), _
                               0, True)
         Call C_display_picture.set_aid_line_start_point(temp_point(2).no, temp_line(0)) '画中点
        End If
   End If
'*******************************************************************************************************
Case "epolygon"
If draw_step = 1 Then
 Call draw_temp_line_for_mouse_up(False)
 If list_type_for_draw = 1 Then
 MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2000, "")
 ElseIf list_type_for_draw = 2 Then
 MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2010, "")
 ElseIf list_type_for_draw = 3 Then
 MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2015, "")
 ElseIf list_type_for_draw = 4 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2020, "")
End If
End If

Case "move_point" 'Then
If yidian_type = 1 Then ' Or ((yidian_type = 9 Or yidian_type = 10 Or yidian_type = 11) And _
    (list_type_for_draw = 3 Or list_type_for_draw = 4 Or list_type_for_draw = 5)) Then
     yidian_type = 2
End If
 If (list_type_for_draw = 3 Or list_type_for_draw = 4 Or list_type_for_draw = 5) And yidian_type >= 1 Then
  If C_curve.set_curve_poi(yidian_no, Int(X), Int(Y), mouse_move_coord.X, mouse_move_coord.Y, 3, yidian_type) Then
   yidian_type = yidian_type + 1
   MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2060, "")
  End If
  Exit Sub
If yidian_no > 0 Then
yidian_step% = 0
Call measur_again
End If
End If
Case "change_picture" 'Then
If draw_step = 1 Then
 If list_type_for_draw = 14 Then
   symmetry_line.p(1).X = Int(X)
   symmetry_line.p(1).Y = Int(Y)
   X = (symmetry_line.p(0).X + symmetry_line.p(1).X) / 2
   Y = (symmetry_line.p(0).Y + symmetry_line.p(1).Y) / 2
   symmetry_line.p(0).X = X + (symmetry_line.p(1).Y - Y)
   symmetry_line.p(0).Y = Y - (symmetry_line.p(1).X - X)
   symmetry_line.p(1).X = 2 * X - symmetry_line.p(0).X
   symmetry_line.p(1).Y = 2 * Y - symmetry_line.p(1).Y
  If set_change_fig = polygon_ Then
     'test Call turn_over_polygon( _
       symmetry_line.p(0), symmetry_line.p(1))
       'test Call draw_change_polygon(1)
  ElseIf set_change_fig = circle_ Then
    'test Call turn_over_circle(symmetry_line.p(0), symmetry_line.p(1))
     'test  Call draw_change_circle(1)
  ElseIf set_change_fig = line_ Then
    'test   Call draw_change_line(5)
    'test Call turn_over_line(symmetry_line.p(0), symmetry_line.p(1))
     'test  Call draw_change_line(5)
          ' Call link_v_for_change_line
  End If
  draw_step = 0
ElseIf list_type_for_draw = 7 Then '轴对称
  If Abs(symmetry_line.p(0).X - symmetry_line.p(1).X) > 4 Or _
      Abs(symmetry_line.p(0).Y - symmetry_line.p(1).Y) > 4 Then
  Call center_display(symmetry_line.p(1)) '画对称轴
   draw_step = 2
 If draw_step = 2 Then
  If set_change_fig = polygon_ Then
    'test Call oxis_symmetry_polygon( _
      Polygon_for_change.p(0), symmetry_line.p(0), symmetry_line.p(1))
      'test Call draw_change_polygon(1)
       'test    Call link_v_for_change_polygon
   Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
          (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
   Call center_display(symmetry_line.p(0))
   Call center_display(symmetry_line.p(1))
  ElseIf set_change_fig = circle_ Then
    ' test Call oxis_symmetry_circle(symmetry_line.p(0), symmetry_line.p(1))
    'test   Call draw_change_circle(1)
    'test       Call link_v_for_change_circle
   Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
          (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
   Call center_display(symmetry_line.p(0))
   Call center_display(symmetry_line.p(1))
     
  ElseIf set_change_fig = line_ Then
     'test  Call draw_change_line(5)
    'test Call oxis_symmetry_line(symmetry_line.p(0), symmetry_line.p(1))
    'test   Call draw_change_line(5)
     'test      Call link_v_for_change_line
  End If
 End If
 Else
  symmetry_line.p(1).X = 0
  symmetry_line.p(1).Y = 0
 End If
 'ElseIf list_type_for_draw = 8 Then
 End If
ElseIf draw_step = 3 Then '动态轴对称
 If list_type_for_draw = 7 Then
  If last_conditions.last_cond(1).change_picture_type = polygon_ Then
  'test  Call oxis_symmetry_polygon( _
      Polygon_for_change.p(0), symmetry_line.p(0), symmetry_line.p(1))
    'test   Call draw_change_polygon(1)
      'test     Call link_v_for_change_polygon
   Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
          (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
   Call center_display(symmetry_line.p(0))
   Call center_display(symmetry_line.p(1))
  'ElseIf last_conditions.last_cond(1).change_picture_type = line_ Then
  ' Call draw_change_polygon(1)
  ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
    'testCall oxis_symmetry_circle(symmetry_line.p(0), symmetry_line.p(1))
     'test  Call draw_change_circle(1)
     'test      Call link_v_for_change_circle
   Draw_form.Line (symmetry_line.p(0).X, symmetry_line.p(0).Y)- _
          (symmetry_line.p(1).X, symmetry_line.p(1).Y), QBColor(12)
   Call center_display(symmetry_line.p(0))
   Call center_display(symmetry_line.p(1))
  End If
   draw_step = 2
ElseIf list_type_for_draw = 8 Then '中心轴对称
  If last_conditions.last_cond(1).change_picture_type = polygon_ Then
    'test Call center_symmetry_polygon( _
      Polygon_for_change.p(0), center_p)
     'test  Call draw_change_polygon(1)
     'test   Call link_v_for_change_polygon
    Call center_display(center_p)
  ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
    'test Call center_symmetry_circle(center_p)
    'test   Call draw_change_circle(1)
    'test    Call link_v_for_change_circle
    Call center_display(center_p)
 ElseIf last_conditions.last_cond(1).change_picture_type = line_ Then
 End If
    draw_step = 2
ElseIf list_type_for_draw = 3 Or list_type_for_draw = 4 Or _
 list_type_for_draw = 5 Then
 '定形平移多边形
  If last_conditions.last_cond(1).change_picture_type = line_ Then
  ElseIf last_conditions.last_cond(1).change_picture_type = polygon_ Then
  'test Call draw_change_polygon(1)
  ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
 'test  Call draw_change_circle(1)
  End If
ElseIf list_type_for_draw = 4 Then
  Draw_form.Line (line_for_move.coord(0).X, _
       line_for_move.coord(0).Y)-(mouse_move_coord.X, mouse_move_coord.Y), QBColor(fill_color)
  Draw_form.Line (line_for_move.coord(0).X, _
       line_for_move.coord(0).Y)-(Int(X), Int(Y)), QBColor(fill_color)
  If Abs(Int(X) - line_for_move.coord(0).X) > 8 Or _
      Abs(Int(Y) - line_for_move.coord(0).Y) > 8 Then
       line_for_move.coord(1).X = Int(X)
        line_for_move.coord(1).Y = Int(Y)
    'Call C_display_picture.m_BPset(Draw_form, Int(x), Int(y), "", 7)
   'Call change_polygon1(move_x&, move_y&)
   operat_is_acting = False
   Else
    Draw_form.Line (line_for_move.coord(0).X, _
       line_for_move.coord(0).Y)-(Int(X), Int(Y)), QBColor(fill_color)
    'Call C_display_picture.m_BPset(Draw_form, line_for_move.coord(0).x, line_for_move.coord(0).y, _
                                    "", 7)
    temp_point(1).no = read_point(mouse_up_coord, 0)
    If temp_point(1).no > 0 Then
     center_p.X = m_poi(temp_point(1).no).data(0).data0.coordinate.X
     center_p.Y = m_poi(temp_point(1).no).data(0).data0.coordinate.Y
    Else
     center_p.X = Int(X)
     center_p.Y = Int(Y)
    End If
    Call center_display(center_p)
    'Call change_polygon2
    operat_is_acting = False
   End If

End If
   operat_is_acting = False
   draw_step = 2
End If
Case "" 'LoadResString_(155) 'Then
 If draw_step = 3 Then
  Draw_form.Line (line_for_move.coord(0).X, _
       line_for_move.coord(0).Y)-(mouse_move_coord.X, mouse_move_coord.Y), QBColor(fill_color)
  Draw_form.Line (line_for_move.coord(0).X, _
       line_for_move.coord(0).Y)-(Int(X), Int(Y)), QBColor(fill_color)
  If Abs(Int(X) - line_for_move.coord(0).X) > 8 Or _
      Abs(Int(Y) - line_for_move.coord(0).Y) > 8 Then
       line_for_move.coord(1).X = Int(X)
        line_for_move.coord(1).Y = Int(Y)
        
  'Call C_display_picture.m_BPset(Draw_form, Int(x), Int(y), "", 7)
   'Call change_polygon1(move_x&, move_y&)
   operat_is_acting = False
   Else
    Draw_form.Line (line_for_move.coord(0).X, _
       line_for_move.coord(0).Y)-(Int(X), Int(Y)), QBColor(fill_color)
    'Call C_display_picture.m_BPset(Draw_form, line_for_move.coord(0).x, line_for_move.coord(0).y, _
                                  "", 7)
    temp_point(1).no = read_point(mouse_up_coord, 0)
    If temp_point(1).no > 0 And Polygon_for_change.p(0).total_v = 2 Then
     If MsgBox(LoadResString_(2065, "\\1\\" + _
       m_poi(temp_point(0).no).data(0).data0.name), 4, "", 0, 0) = vbYes Then
    'test  Call choce_polygon_for_change(temp_point(1))
       draw_step = 1
        Exit Sub
     End If
    End If
    If temp_point(1).no > 0 Then
     center_p.X = m_poi(temp_point(1).no).data(0).data0.coordinate.X
     center_p.Y = m_poi(temp_point(1).no).data(0).data0.coordinate.Y
    Else
     center_p.X = Int(X)
     center_p.Y = Int(Y)
    End If
    Call center_display(center_p)
    'Call change_polygon2
    operat_is_acting = False
   End If
   End If
Case "change_picture" '选择多边形
 If draw_step = 2 Then
  If Abs(init_p.X - Int(X)) > 8 Or Abs(init_p.Y - Int(Y)) > 8 Then
   MDIForm1.StatusBar1.Panels(1).text = ""
         r! = sqr((Int(X) - init_p.X) ^ 2 + (Int(Y) - init_p.Y) ^ 2)
      yidian_stop = False
    Call turn_polygon((Int(X) - init_p.X) / r!, _
     (Int(Y) - init_p.Y) / r!, 0, 0)
      operat_is_acting = False
  Else
   temp_point(1).no = read_point(mouse_up_coord, 0)
   If temp_point(1).no > 0 And Polygon_for_change.p(0).total_v = 2 Then
   'test Call choce_polygon_for_change(temp_point(1))
     draw_step = 1
   End If
  End If
End If
Case "set" 'Then
If list_type_for_draw = 2 And draw_step = 2 Then
        MDIForm1.Text1.visible = True
         MDIForm1.Text2.visible = True
        MDIForm1.Text1.text = LoadResString_(870, "\\1\\" + _
            m_poi(length_(0).poi(0)).data(0).data0.name + _
            m_poi(length_(0).poi(1)).data(0).data0.name)
         MDIForm1.Text2.text = ""
      input_text_statue = True
      MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2165, "")
       MDIForm1.Text2.SetFocus
        measur_step = 0
        'Exit Sub
ElseIf list_type_for_draw = 3 And draw_step = 3 Then
        MDIForm1.Text1.visible = True
         MDIForm1.Text2.visible = True
        MDIForm1.Text1.text = LoadResString_(2175, _
             "\\1\\" + m_poi(temp_point(0).no).data(0).data0.name + _
             "\\2\\" + m_poi(temp_point(1).no).data(0).data0.name + _
               m_poi(temp_point(2).no).data(0).data0.name)
         MDIForm1.Text2.text = ""
      input_text_statue = True
      MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2165, "")
       MDIForm1.Text2.SetFocus
        draw_step = 0
ElseIf list_type_for_draw = 4 Then
        MDIForm1.Text1.visible = True
         MDIForm1.Text2.visible = True
        MDIForm1.Text1.text = LoadResString_(2170, "")
         MDIForm1.Text2.text = ""
      input_text_statue = True
      MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2175, "")
       MDIForm1.Text2.SetFocus
        draw_step = 0
End If
End Select
End If
End Sub

Private Sub set_temp_line_for_draw(ByVal p1%, ByVal k1&, ByVal k2&, _
          p_coord As POINTAPI, p2%, l%)
'用于画垂直，平行，切线
     p2% = 0
     If k1& <> 0 Or k2& <> 0 Then
     t_coord.X = m_poi(p1%).data(0).data0.coordinate.X + k1&
     t_coord.Y = m_poi(p1%).data(0).data0.coordinate.Y + k2&
     l% = line_number(p1%, 0, m_poi(p1%).data(0).data0.coordinate, t_coord, _
                      depend_condition(point_, p1%), _
                      depend_condition(line_, l%), _
                      aid_condition, fill_color, 0, 0)
     End If
End Sub
Public Sub draw_tangent_circle(ByVal tangent_circle_no%, Optional is_delete As Boolean = False)
 Draw_form.Circle (m_tangent_circle(tangent_circle_no%).data(0).data0(1).circle_center.X, _
                     m_tangent_circle(tangent_circle_no%).data(0).data0(1).circle_center.Y), _
                    m_tangent_circle(tangent_circle_no%).data(0).data0(1).circle_radii, _
                        QBColor(fill_color)
   Call m_BPset(Draw_form, m_tangent_circle(tangent_circle_no%).data(0).data0(1).tangent_coord(0), "", fill_color)
   Call m_BPset(Draw_form, m_tangent_circle(tangent_circle_no%).data(0).data0(1).tangent_coord(1), "", fill_color)
 If is_delete = False Then
     m_tangent_circle(tangent_circle_no%).data(0).data0(1) = m_tangent_circle(tangent_circle_no%).data(0).data0(0)
   Draw_form.Circle (m_tangent_circle(tangent_circle_no%).data(0).data0(1).circle_center.X, _
                     m_tangent_circle(tangent_circle_no%).data(0).data0(1).circle_center.Y), _
                    m_tangent_circle(tangent_circle_no%).data(0).data0(1).circle_radii, _
                        QBColor(fill_color)
   Call m_BPset(Draw_form, m_tangent_circle(tangent_circle_no%).data(0).data0(1).tangent_coord(0), "", fill_color)
   Call m_BPset(Draw_form, m_tangent_circle(tangent_circle_no%).data(0).data0(1).tangent_coord(1), "", fill_color)
 End If
End Sub

Private Sub center_display(p As POINTAPI) '显示中心+
 If p.X <> -1000 Or p.Y <> -1000 Then
 Draw_form.Line (p.X - 10, p.Y - 10)-(p.X + 10, p.Y + 10), _
   QBColor(12)
 Draw_form.Line (p.X + 10, p.Y - 10)-(p.X - 10, p.Y + 10), _
   QBColor(12)
 End If
End Sub
Private Sub turn_polygon(ByVal X&, ByVal Y&, ByVal v As Byte, ByVal EX As Byte)
Dim v2!, v3!
Dim t!, i%
Dim m!
 If v = 1 Then
 v2! = 0.02
 move_direct = 1
 ElseIf v = 2 Then
 v2! = -0.02
 move_direct = 2
 Else
 If move_direct = 1 Then
 v2! = 0.02
 ElseIf move_direct = 2 Then
 v2! = -0.02
 Else
 v2! = 0
 End If
 End If
 If EX = 1 Then
  v3! = -0.005
 ElseIf EX = 2 Then
 v3! = 0.005
 Else
 v3! = 0
 End If
MDIForm1.Timer1.Enabled = True
Do
Do
  DoEvents
Loop Until time_act = True
'Call save_draw_data
If Polygon_for_change.similar_ratio > 0.001 Then
Polygon_for_change.similar_ratio = Polygon_for_change.similar_ratio + v3!
End If
t_coord.X = X&
t_coord.Y = Y&
Call change_polygon(t_coord, Polygon_for_change.similar_ratio, True)
time_act = False
Loop Until yidian_stop = True
yidian_stop = False
MDIForm1.Timer1.Enabled = False
End Sub

Private Sub display_sub_menu(X As Single, Y As Single)
 Call C_display_picture.picture_scale
 If X < C_display_picture.picture_top_left_x Or X > C_display_picture.picture_bottom_right_x Or _
      Y < C_display_picture.picture_top_left_y Or Y > C_display_picture.picture_bottom_right_y Then
 '显示子菜单
 Select Case operator
 Case "draw_point_and_line" 'LoadResString_(115)
 Draw_form.PopupMenu MDIForm1.porine, 0, X + 5, Y + 5
 Case "draw_circle"
 Draw_form.PopupMenu MDIForm1.draw_circle, 0, X + 5, Y + 5
 Case "paral_verti", "verti_midpoint"
 Draw_form.PopupMenu MDIForm1.paralandverti0, 0, X + 5, Y + 5
 Case "epolygon"
 Draw_form.PopupMenu MDIForm1.E_polygon, 0, X + 5, Y + 5
 Case "move_point", "set_view_point"
 Draw_form.PopupMenu MDIForm1.move_picture, 0, X + 5, Y + 5
 Case "change_picture"
 Draw_form.PopupMenu MDIForm1.change_picture, 0, X + 5, Y + 5
 Case "measure"
 Draw_form.PopupMenu MDIForm1.mea_and_cal, 0, X + 5, Y + 5
 Case "set"
 Draw_form.PopupMenu MDIForm1.set_for_measure, 0, X + 5, Y + 5
 Case "set_function_data"
 Draw_form.PopupMenu MDIForm1.measure, 0, X + 5, Y + 5
 Case "re_name"
 Draw_form.PopupMenu MDIForm1.re_name, X + 5, Y + 5
 End Select
End If

End Sub
Public Function draw_new_point(input_coord As POINTAPI, ele1 As condition_type, _
                                 ele2 As condition_type, draw_red As Boolean, _
                                  is_set_input As Boolean, need_control As Byte, _
                                   Optional is_no_need_pre_input As Boolean = False) As Integer
                'is_old_point=0 读出所有的点 如果 ele_no%<>0,ele_ty<>0,=1 读出已知点,=-1 ,只读新点
                                           'ele_ty=point_ ele_no=0 读出所有的点,
                           'ele_ty=point_ ele_no=-1 只读新点,
                           'need_control =0 无控制
                           'need_control =1 必需落在已知点
                           'need_control =2 必需落在已知线
                           'need_control =4 必需落在已知圆
                           'need_control=5 不读单点线
                           'need_control =255 有控制
Dim t_coord As POINTAPI
Dim tangent_line_type As Integer
Dim point_type%
't_coord = input_coord
     point_type% = read_inter_point(input_coord, ele1, ele2, temp_point(draw_step).no, _
                           is_set_input, tangent_line_type, need_control)
     draw_new_point = point_type%
 '                        input_coord = t_coord
      If point_type% > 0 Then
       If point_type% = set_new_wenti_cond Then '输入的点已有,但变动位置,切线
         Call init_draw_data
         draw_new_point = True
       End If
       Call set_point_no_reduce(temp_point(draw_step).no, False)
       If draw_red Then
        Call C_display_picture.draw_red_point(temp_point(draw_step).no)
       End If
        'draw_new_point = True '读出已知点
               If is_set_input Then
                Call from_draw_to_input(point_type%, temp_point(draw_step).no, ele1, ele2, tangent_line_type, is_no_need_pre_input)    '新点 引入的输入语句
               End If
      End If
        '选中已知点, 返回点号
End Function

Public Sub plane_geometry_key_press(KeyAscii As Integer)
Dim c$
Dim v$
Dim i%
v$ = ""
c$ = Chr(KeyAscii) '键码
If event_statue = wait_for_modify_char Or _
     event_statue = wait_for_input_char Then
      Call C_display_wenti.m_input_char(Wenti_form.Picture1, c$)
     Wenti_form.SetFocus
Else
c$ = UCase(c$) '大写用于点的名称
If operator = "re_name" Then '重命名
If c$ >= "A" And c$ <= "Z" Then
If choose_point > -1 Then '现实的点
If is_used_char0(c$) Then '是否是已用的点
 MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1615, "")
 Exit Sub
End If
'**********************************************
If list_type_for_draw = 1 Then '重命名所有点
 Call set_point_name(choose_point, c$) 'choose_point 重命名
'  yidian_stop = True
'  For i% = 1 To last_conditions.last_cond(1).point_no
'   If m_poi(i%).data(0).data0.name = "" And m_poi(i%).data(0).data0.visible > 0 Then
'     choose_point = i%
'      yidian_stop = False
'       Call C_display_picture.flash_point(choose_point)
'   Exit Sub
'End If
'Next i%
ElseIf list_type_for_draw = 2 Then '重命名指定点
Call set_point_name(choose_point, c$)
operator = ""
list_type_for_draw = 0
yidian_stop = True
End If
keypress_mark1:
re_name_ty = 0
choose_point = -1
yidian_stop = True
list_type_for_draw = 0
End If
End If
Else
If input_statue_from_p = 1 Then
 temp_key = KeyAscii
input_statue_from_p = 0
End If
If event_statue = wait_for_draw_point Then
 event_statue = input_char_in_draw
  temp_key = KeyAscii
End If
If event_statue = set_measur Then
If list_type_for_draw = 2 Then
keypress_mark:
event_statue = wait_for_set_measur
While event_statue = wait_for_set_measur
  DoEvents
Wend
If Asc(c$) > 45 And Asc(c$) < 58 Then
v$ = v$ + c$
      Call display_m_string(last_measur_string, delete)
       Measur_string(0) = LoadResString_(870, "\\1\\" + _
        m_poi(length_(last_length).poi(0)).data(0).data0.name + _
          m_poi(length_(last_length).poi(1)).data(0).data0.name + "=" + v$ + "_")
      Call display_m_string(last_measur_string, display)
       event_statue = wait_for_set_measur
 GoTo keypress_mark
ElseIf Asc(c$) = 13 Then
      Call display_m_string(last_measur_string, delete)
       Measur_string(0) = LoadResString_(870, "\\1\\" + _
        m_poi(length_(last_length).poi(0)).data(0).data0.name + _
          m_poi(length_(last_length).poi(1)).data(0).data0.name + _
             v$)
      Call display_m_string(last_measur_string, display)
       event_statue = wait_for_set_measur
End If
End If
End If
End If
End If
End Sub
Public Function read_special_point(ByVal point_ty As Integer, _
                  ele1 As condition_type, ele2 As condition_type, _
                   ByVal data_ty As Byte, _
                   ByVal data_no As Integer, ByVal new_point%) As Boolean
Dim input_data_ty As Integer
If data_ty = circle_ Then
        If point_ty% = new_point_on_circle Then
           If ele1.no = data_no Then
             Call C_display_wenti.delete_wenti(7, new_point, _
                  ele1.no, ele2.no, 0)
              read_special_point = True
               Exit Function
           Else
            input_data_ty = 7
           End If
        ElseIf point_ty% = new_point_on_line_circle _
            Or point_ty = new_point_on_line_circle12 _
             Or point_ty = new_point_on_line_circle21 _
              Or point_ty = new_point_on_circle_circle Then
           If ele2.no = data_no Then
              read_special_point = True
               Exit Function
           Else
            input_data_ty = 11
           End If
        ElseIf point_ty = new_point_on_circle_circle12 _
                Or point_ty = new_point_on_circle_circle21 Then
           If ele1.no = data_no Or ele2.no = data_no Then
              read_special_point = True
               Exit Function
           Else
            input_data_ty = 13
           End If
        End If
        If point_ty <> exist_point Then
           Call remove_point(new_point%, True, 1)
           Call C_display_wenti.delete_wenti(input_data_ty, new_point%, ele1.no, ele2.no, 0)
        End If
ElseIf data_ty = line_ Then
        If point_ty = new_point_on_line Then
           If ele1.no = data_no Then
              read_special_point = True
               Exit Function
           Else
            input_data_ty = 1
           End If
        ElseIf point_ty = new_point_on_line_circle _
            Or point_ty = new_point_on_line_circle12 _
             Or point_ty = new_point_on_line_circle21 _
              Or point_ty = new_point_on_circle_circle Then
           If ele1.no = data_no Then
              read_special_point = True
               Exit Function
           Else
            input_data_ty = 11
           End If
        ElseIf point_ty = interset_point_line_line Then
           If ele1.no = data_no Then
              read_special_point = True
               Exit Function
           Else
            input_data_ty = 9
           End If
        End If
        If point_ty <> exist_point Then
           Call remove_point(new_point%, True, 1)
           Call C_display_wenti.delete_wenti(input_data_ty, new_point%, ele1.no, ele2.no, 0)
        End If
End If
End Function
Public Sub init_draw_data()
Dim i%, j%, k%
Dim cir As circle_type
'If operator <> "draw_point_and_line" Or list_type_for_draw <> 1 Then '画点和线的操作可以继续
'Else
'   operator = ""
'End If
'If m_temp_line_for_input.is_using = False Then
temp_tangent_circle_no(0) = 0
temp_tangent_circle_no(1) = 0
m_temp_circle_for_input.data(0) = m_Circ(0).data(0).data0
'Call draw_temp_circle_for_input
'm_temp_line_for_input.data(0) = m_lin(0).data(0).data0
'Call draw_temp_line_for_input
m_temp_line_for_input.is_using = False
'End If
'operator = "draw_point_and_line"
'list_type_for_draw = 1
draw_step = -1
paral_or_verti = 0
draw_operate = False
Call delete_control_data
'draw_point_no = 0
draw_line_no = 0
draw_circle_no = 0
Up_Enabled = False
Move_Enabled = False
yidian_type = 0
measur_step = 0
Draw_form.List1.visible = False
Call C_display_picture.redraw
For i% = 0 To 15
If aid_point(i%) > 0 Then
Call remove_point(aid_point(i%), True, 0)
End If
aid_point(i%) = 0
red_line(i%) = 0
temp_point(i%).no = 0
Next i%
For i% = 0 To 7
temp_circle(i%) = 0
temp_circle(i%) = 0
temp_line(i%) = 0
Next i%
last_m_aid_line = 0
last_aid_point = 0
'For i% = 1 To last_conditions.last_cond(1).line_no
'   If m_lin(i%).data(0).is_change Then
'    m_lin(i%).data(0).data0.type = condition
'     For j% = 1 To m_lin(i%).data(0).data0.in_point(0) - 1
'       m_lin(i%).data(0).data0.color(j%) = condition_color
'     Next j%
'     Call C_display_picture.draw_line(i%, 0, 0)
'   End If
'Next i%
'For i% = 1 To last_conditions.last_cond(1).circle_no
'      Call C_display_picture.redraw_circle(i%)
'Next i%
For i% = 1 To last_conditions.last_cond(1).point_no
   j% = 1
    Do While j% <= m_poi(i%).data(0).in_line(0)
     If m_poi(i%).data(0).in_line(j%) <= 0 Then
      m_poi(i%).data(0).in_line(0) = m_poi(i%).data(0).in_line(0) - 1
       For k% = j% To m_poi(i%).data(0).in_line(0)
        m_poi(i%).data(0).in_line(k%) = m_poi(i%).data(0).in_line(k% + 1)
       Next k%
     Else
       j% = j% + 1
     End If
    Loop
    If m_poi(i%).data(0).data0.color = conclusion_color Then
       m_poi(i%).data(0).data0.color = condition_color
       Call C_display_picture.set_m_point_color(i%, condition_color)
    End If
Next i%
'Call C_display_picture.redraw_point(0)
'Call C_display_picture.redraw_circle(0)
End Sub
Private Function is_special_point(point_coord As POINTAPI) As Boolean
Dim i%
   If control_data.select_point(0) = 0 Then
       is_special_point = True
         GoTo mark1
   Else
    For i% = 1 To control_data.select_point(0)
        If is_same_POINTAPI(point_coord, m_poi(control_data.forbid_point(i%)).data(0).data0.coordinate) Then
         is_special_point = True
          GoTo mark1
        End If
    Next i%
    is_special_point = False
     Exit Function
   End If
mark1:
   If control_data.forbid_line(0) = 0 Then
        is_special_point = True
         GoTo mark2
   Else
    For i% = 1 To control_data.forbid_line(0)
        If is_same_POINTAPI(point_coord, m_poi(control_data.forbid_point(i%)).data(0).data0.coordinate) Then
            is_special_point = False
             Exit Function
         End If
    Next i%
       is_special_point = True
mark2:
   End If
'
End Function

Private Function is_special_line_or_circle(ele As condition_type) As Boolean
'判定中ele1 ,ele2是否含有s_ele is_special_geometry_ele=0,true 确定,false否定
Dim i%
'**********************************************************************************
If ele.ty = 0 Or ele.no = 0 Then
   If control_data.select_line(0) > 0 Or control_data.select_circle(0) > 0 Then
       is_special_line_or_circle = False
   Else
       is_special_line_or_circle = True
   End If
       Exit Function
ElseIf ele.ty = line_ Then
   If control_data.select_line(0) = 0 Then
       is_special_line_or_circle = True
         GoTo mark1
   Else
    For i% = 1 To control_data.select_line(0)
        If ele.no = control_data.select_line(i%) Then
         is_special_line_or_circle = True
          GoTo mark1
        End If
    Next i%
    is_special_line_or_circle = False
     Exit Function
   End If
mark1:
   If control_data.forbid_line(0) = 0 Then
        is_special_line_or_circle = True
         GoTo mark2
   Else
    For i% = 1 To control_data.forbid_line(0)
        If ele.no = control_data.forbid_line(i%) Then
            is_special_line_or_circle = False
             Exit Function
         End If
    Next i%
       is_special_line_or_circle = True
mark2:
   End If
'*******************************************************************************
ElseIf ele.ty = circle_ Then
   If control_data.select_circle(0) = 0 Then
        is_special_line_or_circle = True
         GoTo mark3
   Else
    For i% = 1 To control_data.select_circle(0)
        If ele.no = control_data.select_circle(i%) Then
         is_special_line_or_circle = True
         GoTo mark3
        End If
    Next i%
    is_special_line_or_circle = False
     Exit Function
   End If
mark3:
   If control_data.forbid_circle(0) = 0 Then
        is_special_line_or_circle = True
        GoTo mark4
   Else
    For i% = 1 To control_data.forbid_circle(0)
        If ele.no = control_data.forbid_circle(i%) Then
            is_special_line_or_circle = False
             Exit Function
        End If
    Next i%
    is_special_line_or_circle = False
mark4:
    End If
Else
          is_special_line_or_circle = True
   
'***********************************************************************
End If
'ElseIf ele.ty = plane_ Then
'   is_special_geometry_ele_ = 1
'   If control_data.select_plane(0) = 0 Then
        'is_special_geometry_ele_ = is_special_geometry_ele_ * 1
'   Else
'    For i% = 1 To control_data.select_plane(0)
'        If ele.no = control_data.select_plane(i%) Then
'         GoTo is_special_geometry_ele_7
'            'is_special_geometry_ele_ = is_special_geometry_ele_ * 1
'        End If
'    Next i%
'    is_special_geometry_ele_ = is_special_geometry_ele_ * 0
'    End If
'is_special_geometry_ele_7:
'   If control_data.forbid_circle(0) = 0 Then
        'is_special_geometry_ele_ = is_special_geometry_ele_ * 1
'   Else
'    For i% = 1 To control_data.forbid_plane(0)
'        If ele.no = control_data.forbid_plane(i%) Then
'            is_special_geometry_ele_ = -1
'         GoTo is_special_geometry_ele_8
'        End If
'    Next i%
    'is_special_geometry_ele_ = is_special_geometry_ele_ * 1
'is_special_geometry_ele_8:
' End If

'Else
' is_special_geometry_ele_ = 1
'End If
End Function
Private Function is_special_geometry_ele(point_coord As POINTAPI, ele1 As condition_type, ele2 As condition_type, _
                                          need_control As Byte) As Boolean
'选中C_ele1.ty,C_ele_no但不是
Dim ty(1) As Integer
If need_control = 0 Then
 is_special_geometry_ele = True
ElseIf need_control = 1 Then
 is_special_geometry_ele = (is_special_line_or_circle(ele1) Or is_special_line_or_circle(ele2)) And _
                       is_special_point(point_coord)
ElseIf need_control = 2 Then
ElseIf need_control = 4 Then
ElseIf need_control = 255 Then
 If ele1.no = 0 And ele2.no = 0 Then
 is_special_geometry_ele = (is_special_line_or_circle(ele1) Or is_special_line_or_circle(ele2)) And _
                       is_special_point(point_coord)
 ElseIf ele1.no = 0 Then
 is_special_geometry_ele = is_special_line_or_circle(ele2) And _
                       is_special_point(point_coord)
 ElseIf ele2.no = 0 Then
 is_special_geometry_ele = is_special_line_or_circle(ele1) And _
                       is_special_point(point_coord)
 Else
 is_special_geometry_ele = (is_special_line_or_circle(ele1) Or is_special_line_or_circle(ele2)) And _
                       is_special_point(point_coord)
 End If
Else
  is_special_geometry_ele = True
  ' End If
 End If
End Function
Public Function depend_condition(ByVal ty As Byte, ByVal no%) As condition_type
      depend_condition.ty = ty
      depend_condition.no = no%
End Function

Public Function draw_vertical_mid_point(ByVal p1%, ByVal p2%, mid_point%) As Integer
Dim t_coord As POINTAPI
Dim l%
Dim ele(1) As condition_type
       MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1425, "")
         draw_vertical_mid_point = get_midpoint(p1%, 0, p2%, 0, 0, 0, 0, 0)
               t_coord = divide_POINTAPI_by_number(add_POINTAPI(m_poi(temp_point(0).no).data(0).data0.coordinate, _
                             m_poi(temp_point(1).no).data(0).data0.coordinate), 2)
              l% = line_number0(p1%, p2%, 0, 0)
                mid_point% = m_point_number(t_coord, condition, 1, condition_color, "", _
                              depend_condition(point_, p1%), depend_condition(point_, p2%), 0, True)
                Call add_point_to_line(mid_point%, l%, 0, True, True, 0)
                draw_vertical_mid_point = line_number(mid_point%, 0, m_poi(mid_point%).data(0).data0.coordinate, _
                                add_POINTAPI(t_coord, verti_POINTAPI(minus_POINTAPI(m_poi(p2%).data(0).data0.coordinate, _
                                  m_poi(p1%).data(0).data0.coordinate))), _
                                   depend_condition(point_, mid_point), _
                                    depend_condition(0, 0), _
                                      aid_condition, fill_color, 0, 0)
         
   End Function
Public Function set_temp_paral_and_vertical_line(ByVal p1%, ByVal line_no%) As Integer
Dim i%, j%
Dim t_coord(1) As POINTAPI
Dim no_has_temp_line(1) As Boolean
     t_coord(0) = add_POINTAPI(m_poi(p1%).data(0).data0.coordinate, minus_POINTAPI( _
          m_poi(m_lin(line_no%).data(0).data0.poi(0)).data(0).data0.coordinate, _
            m_poi(m_lin(line_no%).data(0).data0.poi(1)).data(0).data0.coordinate))
     t_coord(1) = add_POINTAPI(m_poi(p1%).data(0).data0.coordinate, verti_POINTAPI(minus_POINTAPI( _
          m_poi(m_lin(line_no%).data(0).data0.poi(0)).data(0).data0.coordinate, _
            m_poi(m_lin(line_no%).data(0).data0.poi(1)).data(0).data0.coordinate)))
If is_point_in_line3(p1%, m_lin(line_no%).data(0).data0, 0) Then
    set_temp_paral_and_vertical_line = line_number(p1%, 0, m_poi(p1%).data(0).data0.coordinate, _
                    t_coord(1), depend_condition(point_, p1%), depend_condition(line_, line_no%), _
                      aid_condition, fill_color, 0, 0)
      Call C_display_picture.set_aid_line_start_point(-1, set_temp_paral_and_vertical_line)
Else 'End If
For i% = 1 To m_poi(p1%).data(0).in_line(0)
     If m_poi(p1%).data(0).in_line(i%) = line_no% Then '点p1%在直线line_no%上
        no_has_temp_line(0) = True
     End If
     If no_has_temp_line(0) = False Then
     For j% = 1 To m_lin(line_no%).data(0).in_paral(0).lin
        If m_lin(line_no%).data(0).in_paral(j%).line_no = m_poi(p1%).data(0).in_line(i%) Then
           no_has_temp_line(1) = True
        End If
     Next j%
     End If
     For j% = 1 To m_lin(line_no%).data(0).in_verti(0).lin
        If m_lin(line_no%).data(0).in_verti(j%).line_no = m_poi(p1%).data(0).in_line(i%) Then
           no_has_temp_line(1) = True
        End If
     Next j%
Next i%
If no_has_temp_line(0) = False And no_has_temp_line(1) = False Then '过p1%没有平行line_no%的直线
  set_temp_paral_and_vertical_line = line_number(p1%, 0, m_poi(p1%).data(0).data0.coordinate, _
                    t_coord(0), depend_condition(point_, p1%), depend_condition(line_, line_no%), _
                      aid_condition, fill_color, 0, 0)
  Call C_display_picture.set_aid_line_start_point(p1%, set_temp_paral_and_vertical_line)
ElseIf no_has_temp_line(1) = False Then
  set_temp_paral_and_vertical_line = line_number(p1%, 0, m_poi(p1%).data(0).data0.coordinate, _
                    t_coord(1), depend_condition(point_, p1%), depend_condition(line_, line_no%), _
                      aid_condition, fill_color, 0, 0)

ElseIf no_has_temp_line(0) = False Then
  set_temp_paral_and_vertical_line = line_number(p1%, 0, m_poi(p1%).data(0).data0.coordinate, _
                    t_coord(0), depend_condition(point_, p1%), depend_condition(line_, line_no%), _
                      aid_condition, fill_color, 0, 0)

       ' Call set_temp_line_for_draw(p1%, _
          m_poi(m_lin(line_no%).data(0).data0.poi(0)).data(0).data0.coordinate.Y - _
            m_poi(m_lin(line_no%).data(0).data0.poi(1)).data(0).data0.coordinate.Y, _
          m_poi(m_lin(line_no%).data(0).data0.poi(1)).data(0).data0.coordinate.X - _
            m_poi(m_lin(line_no%).data(0).data0.poi(0)).data(0).data0.coordinate.X, _
             pointapi0, temp_point(7).no, temp_line(7))
End If
End If
End Function
Public Sub set_select_point(Optional p1 As Integer = 0, _
                             Optional p2 As Integer = 0, _
                             Optional p3 As Integer = 0, _
                             Optional p4 As Integer = 0)
control_data.select_point(0) = 0
If p1 > 0 Then
control_data.select_point(0) = control_data.select_point(0) + 1
control_data.select_point(control_data.select_point(0)) = p1%
End If
If p2 > 0 Then
control_data.select_point(0) = control_data.select_point(0) + 1
control_data.select_point(control_data.select_point(0)) = p2%
End If
If p3 > 0 Then
control_data.select_point(0) = control_data.select_point(0) + 1
control_data.select_point(control_data.select_point(0)) = p3%
End If
If p4 > 0 Then
control_data.select_point(0) = control_data.select_point(0) + 1
control_data.select_point(control_data.select_point(0)) = p4%
End If
End Sub
Public Sub set_forbid_point(Optional p1 As Integer = 0, _
                             Optional p2 As Integer = 0, _
                             Optional p3 As Integer = 0, _
                             Optional p4 As Integer = 0)
control_data.forbid_point(0) = 0
If p1 > 0 Then
control_data.forbid_point(0) = control_data.forbid_point(0) + 1
control_data.forbid_point(control_data.forbid_point(0)) = p1%
End If
If p2 > 0 Then
control_data.forbid_point(0) = control_data.forbid_point(0) + 1
control_data.forbid_point(control_data.forbid_point(0)) = p2%
End If
If p3 > 0 Then
control_data.forbid_point(0) = control_data.forbid_point(0) + 1
control_data.forbid_point(control_data.forbid_point(0)) = p3%
End If
If p4 > 0 Then
control_data.forbid_point(0) = control_data.forbid_point(0) + 1
control_data.forbid_point(control_data.forbid_point(0)) = p4%
End If
End Sub
Public Sub set_select_line(Optional l1 As Integer = 0, _
                             Optional l2 As Integer = 0, _
                             Optional l3 As Integer = 0, _
                             Optional l4 As Integer = 0)
control_data.select_line(0) = 0
If l1 > 0 Then
control_data.select_line(0) = control_data.select_line(0) + 1
control_data.select_line(control_data.select_line(0)) = l1%
End If
If l2 > 0 Then
control_data.select_line(0) = control_data.select_line(0) + 1
control_data.select_line(control_data.select_line(0)) = l2%
End If
If l3 > 0 Then
control_data.select_line(0) = control_data.select_line(0) + 1
control_data.select_line(control_data.select_line(0)) = l3%
End If
If l4 > 0 Then
control_data.select_line(0) = control_data.select_line(0) + 1
control_data.select_line(control_data.select_line(0)) = l4%
End If
End Sub
Public Sub set_forbid_line(Optional l1 As Integer = 0, _
                             Optional l2 As Integer = 0, _
                             Optional l3 As Integer = 0, _
                             Optional l4 As Integer = 0)
control_data.forbid_line(0) = 0
If l1 > 0 Then
control_data.forbid_line(0) = control_data.forbid_line(0) + 1
control_data.forbid_line(control_data.forbid_line(0)) = l1%
End If
If l2 > 0 Then
control_data.forbid_line(0) = control_data.forbid_line(0) + 1
control_data.forbid_line(control_data.forbid_line(0)) = l2%
End If
If l3 > 0 Then
control_data.forbid_line(0) = control_data.forbid_line(0) + 1
control_data.forbid_line(control_data.forbid_line(0)) = l3%
End If
If l4 > 0 Then
control_data.forbid_line(0) = control_data.forbid_line(0) + 1
control_data.forbid_line(control_data.forbid_line(0)) = l4%
End If
End Sub
Public Sub delete_control_data()
 control_data.forbid_circle(0) = 0
 control_data.forbid_line(0) = 0
 control_data.forbid_plane(0) = 0
 control_data.forbid_point(0) = 0
 control_data.select_circle(0) = 0
 control_data.select_line(0) = 0
 control_data.select_plane(0) = 0
 control_data.select_point(0) = 0
End Sub
Public Sub set_select_circle(Optional c1 As Integer = 0, _
                             Optional c2 As Integer = 0, _
                             Optional c3 As Integer = 0, _
                             Optional C4 As Integer = 0)
control_data.select_circle(0) = 0
If c1 > 0 Then
control_data.select_circle(0) = control_data.select_circle(0) + 1
control_data.select_circle(control_data.select_circle(0)) = c1%
End If
If c2 > 0 Then
control_data.select_circle(0) = control_data.select_circle(0) + 1
control_data.select_circle(control_data.select_circle(0)) = c2%
End If
If c3 > 0 Then
control_data.select_circle(0) = control_data.select_circle(0) + 1
control_data.select_circle(control_data.select_circle(0)) = c3%
End If
If C4 > 0 Then
control_data.select_circle(0) = control_data.select_circle(0) + 1
control_data.select_circle(control_data.select_circle(0)) = C4%
End If
End Sub
Public Sub set_forbid_circle(Optional c1 As Integer = 0, _
                             Optional c2 As Integer = 0, _
                             Optional c3 As Integer = 0, _
                             Optional C4 As Integer = 0)
control_data.forbid_circle(0) = 0
If c1 > 0 Then
control_data.forbid_circle(0) = control_data.forbid_circle(0) + 1
control_data.forbid_circle(control_data.forbid_circle(0)) = c1%
End If
If c2 > 0 Then
control_data.forbid_circle(0) = control_data.forbid_circle(0) + 1
control_data.forbid_circle(control_data.forbid_circle(0)) = c2%
End If
If c3 > 0 Then
control_data.forbid_circle(0) = control_data.forbid_circle(0) + 1
control_data.forbid_circle(control_data.forbid_circle(0)) = c3%
End If
If C4 > 0 Then
control_data.forbid_circle(0) = control_data.forbid_circle(0) + 1
control_data.forbid_circle(control_data.forbid_circle(0)) = C4%
End If
End Sub
Public Sub set_select_plane(Optional p1 As Integer = 0, _
                             Optional p2 As Integer = 0, _
                             Optional p3 As Integer = 0, _
                             Optional p4 As Integer = 0)
control_data.select_plane(0) = 0
If p1 > 0 Then
control_data.select_plane(0) = control_data.select_plane(0) + 1
control_data.select_plane(control_data.select_plane(0)) = p1%
End If
If p2 > 0 Then
control_data.select_plane(0) = control_data.select_plane(0) + 1
control_data.select_plane(control_data.select_plane(0)) = p2%
End If
If p3 > 0 Then
control_data.select_plane(0) = control_data.select_plane(0) + 1
control_data.select_plane(control_data.select_plane(0)) = p3%
End If
If p4 > 0 Then
control_data.select_plane(0) = control_data.select_plane(0) + 1
control_data.select_plane(control_data.select_plane(0)) = p4%
End If
End Sub
Public Sub set_forbid_plane(Optional p1 As Integer = 0, _
                             Optional p2 As Integer = 0, _
                             Optional p3 As Integer = 0, _
                             Optional p4 As Integer = 0)
control_data.forbid_plane(0) = 0
If p1 > 0 Then
control_data.forbid_plane(0) = control_data.forbid_plane(0) + 1
control_data.forbid_plane(control_data.forbid_plane(0)) = p1%
End If
If p2 > 0 Then
control_data.forbid_plane(0) = control_data.forbid_plane(0) + 1
control_data.forbid_plane(control_data.forbid_plane(0)) = p2%
End If
If p3 > 0 Then
control_data.forbid_plane(0) = control_data.forbid_plane(0) + 1
control_data.forbid_plane(control_data.forbid_plane(0)) = p3%
End If
If p4 > 0 Then
control_data.forbid_plane(0) = control_data.forbid_plane(0) + 1
control_data.forbid_plane(control_data.forbid_plane(0)) = p4%
End If
End Sub



Public Sub set_tangent_line_data(p_coord1 As POINTAPI, p_coord2 As POINTAPI, circ1 As Integer, circ2 As Integer, ele1 As condition_type, ele2 As condition_type, _
                          visible As Byte, tangent_line_ty As Integer)
Dim i%
For i% = last_conditions.last_cond(1).tangent_line_no + 1 To last_conditions.last_cond(0).tangent_line_no  '已有数据和未进入数据库的切线
     If is_same_POINTAPI(p_coord1, tangent_line(i%).data(0).coordinate(0)) And _
           is_same_POINTAPI(p_coord2, tangent_line(i%).data(0).coordinate(1)) Then
           If ele1.ty = tangent_line(i%).data(0).ele(0).ty And ele1.no = tangent_line(i%).data(0).ele(0).no And _
               ele2.ty = tangent_line(i%).data(0).ele(1).ty And ele2.no = tangent_line(i%).data(0).ele(1).no Then
                Exit Sub
           End If
     ElseIf is_same_POINTAPI(p_coord1, tangent_line(i%).data(0).coordinate(1)) And _
              is_same_POINTAPI(p_coord2, tangent_line(i%).data(0).coordinate(0)) Then
            If ele2.ty = tangent_line(i%).data(0).ele(0).ty And ele2.no = tangent_line(i%).data(0).ele(0).no And _
               ele1.ty = tangent_line(i%).data(0).ele(1).ty And ele1.no = tangent_line(i%).data(0).ele(1).no Then
                Exit Sub
           End If
     End If
Next i%
If last_conditions.last_cond(0).tangent_line_no Mod 10 = 0 Then
    ReDim Preserve tangent_line(last_conditions.last_cond(0).tangent_line_no + 10) As tangent_line_type
End If
   last_conditions.last_cond(0).tangent_line_no = last_conditions.last_cond(0).tangent_line_no + 1
    tangent_line(last_conditions.last_cond(0).tangent_line_no).data(0).coordinate(0) = p_coord1 '切线端点坐标3
    tangent_line(last_conditions.last_cond(0).tangent_line_no).data(0).coordinate(1) = p_coord2
    tangent_line(last_conditions.last_cond(0).tangent_line_no).data(0).circ(0) = circ1 '相切圆
    tangent_line(last_conditions.last_cond(0).tangent_line_no).data(0).circ(1) = circ2
    tangent_line(last_conditions.last_cond(0).tangent_line_no).data(0).ele(0) = ele1 '确定切线的元素
    tangent_line(last_conditions.last_cond(0).tangent_line_no).data(0).ele(1) = ele2
    tangent_line(last_conditions.last_cond(0).tangent_line_no).data(0).visible = visible
    tangent_line(last_conditions.last_cond(0).tangent_line_no).tangent_type = tangent_line_ty '切线的类型
    Call draw_tangent_line(last_conditions.last_cond(0).tangent_line_no, 1)
End Sub

Private Sub read_circles_and_lines_from_point(in_point%)
Dim i%, j%
For i% = 0 To 7
If m_poi(in_point%).data(0).in_circle(i%) > 0 Then
 temp_circles_for_draw(i%) = m_poi(in_point%).data(0).in_circle(i%)
End If
If m_poi(in_point%).data(0).in_line(i%) > 0 Then
  temp_lines_for_draw(i%) = m_poi(in_point%).data(0).in_line(i%)
End If
Next i%
For i% = 1 To last_conditions.last_cond(1).circle_no
   If m_Circ(i%).data(0).data0.center = in_point% Then
        temp_circles_for_draw(0) = temp_circles_for_draw(0) + 1
         temp_circles_for_draw(temp_circles_for_draw(0)) = i%
   End If
Next i%
End Sub

Public Sub set_tangent_line_from_point_to_circle(ByVal in_circle_no%, _
                                          ByVal in_point_no%, ByVal no_reduce As Byte)

Dim i%
Dim sr&
'Dim ty As Byte
Dim p_coord(1) As POINTAPI
'Dim p As POINTAPI
If no_reduce = 255 Then
 Exit Sub
End If
'*****************************************************************************************************
 For i% = 1 To m_Circ(in_circle_no%).data(0).data0.in_point(0) '输入的点是否在圆上
  If m_Circ(in_circle_no%).data(0).data0.in_point(i%) = in_point_no% Then
  p_coord(0) = m_poi(in_point_no%).data(0).data0.coordinate '设立过圆上一点作切线
  p_coord(1) = add_POINTAPI(p_coord(0), verti_POINTAPI(minus_POINTAPI( _
                  m_Circ(in_circle_no%).data(0).data0.c_coord, p_coord(0))))
                Call set_tangent_line_data(p_coord(0), p_coord(1), in_circle_no%, 0, _
                depend_condition(circle_, in_circle_no%), depend_condition(point_, in_point_no%), 4, tangent_line_by_point_on_circle)  '单端点切线
                 Exit Sub
  End If
 Next i%
 If m_Circ(in_circle_no%).data(0).data0.center > 0 Then
         m_Circ(in_circle_no%).data(0).data0.c_coord = m_poi(m_Circ(in_circle_no%).data(0).data0.center).data(0).data0.coordinate
 End If
        sr& = distance_of_two_POINTAPI(m_Circ(in_circle_no%).data(0).data0.c_coord, m_poi(in_point_no%).data(0).data0.coordinate)
        Call inter_point_circle_circle_by_pointapi(m_Circ(in_circle_no%).data(0).data0.c_coord, _
              m_Circ(in_circle_no%).data(0).data0.radii, mid_POINTAPI(m_Circ(in_circle_no%).data(0).data0.c_coord, _
               m_poi(in_point_no%).data(0).data0.coordinate), sr& / 2, p_coord(0), p_coord(1))
  Call set_tangent_line_data(m_poi(in_point_no%).data(0).data0.coordinate, p_coord(0), 0, in_circle_no%, _
                 depend_condition(point_, in_point_no%), depend_condition(circle_, in_circle_no%), 1, tangent_line_by_point_out_off_circle12)
  Call set_tangent_line_data(m_poi(in_point_no%).data(0).data0.coordinate, p_coord(1), 0, in_circle_no%, _
                 depend_condition(point_, in_point_no%), depend_condition(circle_, in_circle_no%), 1, tangent_line_by_point_out_off_circle21)
End Sub

Private Sub draw_tangent_line_by_ty(n%)
If tangent_line(n%).data(0).visible = 2 Then
     Draw_form.Line (tangent_line(n%).data(0).old_coordinate(0).X, tangent_line(n%).data(0).old_coordinate(0).Y)- _
      (tangent_line(n%).data(0).old_coordinate(1).X, tangent_line(n%).data(0).old_coordinate(1).Y), QBColor(fill_color)
ElseIf tangent_line(n%).data(0).visible >= 4 Then
      Draw_form.Line (tangent_line(n%).data(0).old_coordinate(0).X, tangent_line(n%).data(0).old_coordinate(0).Y)- _
      (tangent_line(n%).data(0).old_coordinate(1).X, tangent_line(n%).data(0).old_coordinate(1).Y), QBColor(fill_color)
Else
   If tangent_line(n%).data(0).old_coordinate(0).X < 10000 And _
        tangent_line(n%).data(0).old_coordinate(0).Y < 10000 Then
      Call m_BPset(Draw_form, tangent_line(n%).data(0).coordinate(0), "", fill_color)
      Call m_BPset(Draw_form, tangent_line(n%).data(0).coordinate(1), "", fill_color)
     Draw_form.Line (tangent_line(n%).data(0).old_coordinate(0).X, tangent_line(n%).data(0).old_coordinate(0).Y)- _
      (tangent_line(n%).data(0).old_coordinate(1).X, tangent_line(n%).data(0).old_coordinate(1).Y), QBColor(fill_color)

   End If
        
End If
End Sub


Public Function set_tangent_circle_data(center As POINTAPI, radii As Long, ByVal tangent_point1%, tangent_point_coord_1 As POINTAPI, _
               ele1 As condition_type, ByVal tangent_point2%, tangent_point_coord_2 As POINTAPI, ele2 As condition_type, _
                tangent_circle_ty As Integer, Optional tangent_circle_no As Integer = 0) As Integer
Dim temp_radii_1(1) As Integer
Dim temp_radii_2(1) As Integer
Dim dis%
Dim t_tangent_point_coord  As POINTAPI
Dim i%
Dim t_ele As condition_type
If ele1.ty < ele2.ty Or (ele1.ty = ele2.ty And ele1.no < ele2.no) Then
   Call exchange_two_integer(tangent_point1%, tangent_point2%)
   t_tangent_point_coord = tangent_point_coord_1
   tangent_point_coord_1 = tangent_point_coord_2
   tangent_point_coord_2 = t_tangent_point_coord
   t_ele = ele1
   ele1 = ele2
   ele2 = t_ele
End If
If tangent_circle_no = 0 Then
For i% = last_conditions.last_cond(1).tangent_circle_no + 1 To last_conditions.last_cond(0).tangent_circle_no  '已有数据和未进入数据库的切线
     If ele1.ty = m_tangent_circle(i%).data(0).ele(0).ty And ele1.no = m_tangent_circle(i%).data(0).ele(0).no And _
               ele2.ty = m_tangent_circle(i%).data(0).ele(1).ty And ele2.no = m_tangent_circle(i%).data(0).ele(1).no And _
                 tangent_circle_ty = m_tangent_circle(i%).data(0).tangent_circle_ty And _
                  m_tangent_circle(i%).data(0).tangent_poi(0) = 0 And m_tangent_circle(i%).data(0).tangent_poi(1) = 0 Then
                                  set_tangent_circle_data = i%
               GoTo set_tangent_circle_data_mark1
     End If
Next i%
If last_conditions.last_cond(0).tangent_circle_no Mod 10 = 0 Then
    ReDim Preserve m_tangent_circle(last_conditions.last_cond(0).tangent_circle_no + 10) As tangent_circle_type
End If
   last_conditions.last_cond(0).tangent_circle_no = last_conditions.last_cond(0).tangent_circle_no + 1
       set_tangent_circle_data = last_conditions.last_cond(0).tangent_circle_no
Else
    set_tangent_circle_data = tangent_circle_no
End If
set_tangent_circle_data_mark1:
    m_tangent_circle(set_tangent_circle_data).data(0).data0(1) = _
                         m_tangent_circle(set_tangent_circle_data).data(0).data0(0)
    '**********************************************************************************************
    m_tangent_circle(set_tangent_circle_data).data(0).data0(0).tangent_coord(0) = tangent_point_coord_1  '切线端点坐标3
    m_tangent_circle(set_tangent_circle_data).data(0).data0(0).tangent_coord(1) = tangent_point_coord_2
    m_tangent_circle(set_tangent_circle_data).data(0).data0(0).visible = 1
    m_tangent_circle(set_tangent_circle_data).data(0).tangent_circle_ty = tangent_circle_ty '切线的类型
    m_tangent_circle(set_tangent_circle_data).data(0).data0(0).circle_radii = radii
    m_tangent_circle(set_tangent_circle_data).data(0).data0(0).circle_center = center
    '**************************************************************************************************
  If tangent_circle_no = 0 Then
    m_tangent_circle(set_tangent_circle_data).data(0).tangent_poi(0) = tangent_point1%  '切点
    m_tangent_circle(set_tangent_circle_data).data(0).tangent_poi(1) = tangent_point2%
    m_tangent_circle(set_tangent_circle_data).data(0).ele(0) = ele1 '确定切线的元素
    m_tangent_circle(set_tangent_circle_data).data(0).ele(1) = ele2
    Call draw_tangent_circle(set_tangent_circle_data)
  Else
    Call draw_tangent_circle(set_tangent_circle_data)
  End If
End Function

Public Sub draw_temp_line_for_input(Optional draw_type = 0)
   If draw_type = 1 Then '消除临时画的直线
      Call Drawline(Draw_form, QBColor(m_temp_line_for_input.data(1).total_color), 0, _
                  m_temp_line_for_input.data(1).end_point_coord(0), _
                   m_temp_line_for_input.data(1).end_point_coord(1), _
                     0, m_temp_line_for_input.data(1).visible)
     m_temp_line_for_input.data(1) = m_temp_line_for_input.data(0)
   Else
   Call Drawline(Draw_form, QBColor(m_temp_line_for_input.data(1).total_color), 0, _
                  m_temp_line_for_input.data(1).end_point_coord(0), _
                   m_temp_line_for_input.data(1).end_point_coord(1), _
                     0, m_temp_line_for_input.data(1).visible)
     m_temp_line_for_input.data(1) = m_temp_line_for_input.data(0)
   Call Drawline(Draw_form, QBColor(m_temp_line_for_input.data(1).total_color), 0, _
                  m_temp_line_for_input.data(1).end_point_coord(0), _
                   m_temp_line_for_input.data(1).end_point_coord(1), _
                     0, m_temp_line_for_input.data(1).visible)
   End If
End Sub

Public Sub draw_temp_circle_for_input()
If m_temp_circle_for_input.is_using Then
 Draw_form.Circle (m_temp_circle_for_input.data(1).c_coord.X, m_temp_circle_for_input.data(1).c_coord.Y), _
                    m_temp_circle_for_input.data(1).radii&, QBColor(m_temp_circle_for_input.data(1).color)
     m_temp_circle_for_input.data(1) = m_temp_circle_for_input.data(0)
 Draw_form.Circle (m_temp_circle_for_input.data(1).c_coord.X, m_temp_circle_for_input.data(1).c_coord.Y), _
                    m_temp_circle_for_input.data(1).radii&, QBColor(m_temp_circle_for_input.data(1).color)
End If
End Sub

Public Sub draw_temp_line_for_mouse_up(is_input_complete As Boolean, _
                   Optional select_point_no As Integer = 0, Optional select_line_no As Integer = 0, _
                     Optional select_circle_no As Integer = 0)
'****************************************************************************************************************
         Call set_select_point(select_point_no)
         Call set_select_line(select_line_no)
         Call set_select_circle(select_circle_no)
         Call set_forbid_point(temp_point(draw_step - 1).no)
         Call set_forbid_line
         Call set_forbid_circle
'****************************************************************************************************************
       If draw_new_point(mouse_up_coord, ele1, ele2, red, True, 255) > 0 Then  '输入坐标点
         If temp_line(draw_line_no) = 0 Then
            temp_line(draw_line_no) = line_number(temp_point(draw_step - 1).no, temp_point(draw_step).no, _
                                  pointapi0, pointapi0, _
                                  depend_condition(point_, temp_point(draw_step - 1).no), _
                                  depend_condition(point_, temp_point(draw_step).no), _
                                  condition, conclusion_color, 1, 0) '建立直线
          Else
            If is_input_complete = False Then '未完成输入
             m_lin(temp_line(draw_line_no)).data(0).data0.color(1) = conclusion_color
            End If
             m_lin(temp_line(draw_line_no)).data(0).is_change = True
              Call C_display_picture.re_draw_line(temp_line(draw_line_no)) '重画直线
               'm_lin(temp_line(draw_line_no)).data(0).is_change = False
          End If
        draw_line_no = draw_line_no + 1 '
       move_init = draw_step
       Up_Enabled = True
       Move_Enabled = False
      Else
       If m_temp_line_for_input.is_using Then
         Call draw_temp_line_for_input(1) '消除临时输入数据
         m_temp_line_for_input.is_using = False
       End If
       draw_step = draw_step - 2 '作直线，失败，退回上一步
     End If
End Sub
Function read_line(ByVal last_line%, in_coord As POINTAPI, point_no%, _
           out_coord As POINTAPI, is_set_data As Boolean, input_ty As Integer, Optional ty As Byte = 0) As Integer
           'input_ty 输入切线的类型
Dim t_in_coord As POINTAPI
t_in_coord = in_coord
'******************************************************
        read_line = read_tangent_line(in_coord, point_no%, out_coord, input_ty, is_set_data)  'And need_control = 5 Then '搜索切线，并建立切线,必须落在切线上
       If read_line = 0 Then
            read_line = C_display_picture.read_line(last_line%, t_in_coord.X, in_coord.Y, point_no%, out_coord.X, out_coord.Y, is_set_data, ty)
       End If
'**************************************************************************************************
End Function
Public Function out_tangent_line_for_two_circle(c1%, c2%)
If m_Circ(c1%).data(0).data0.radii < m_Circ(c2%).data(0).data0.radii < -3 Then

ElseIf m_Circ(c1%).data(0).data0.radii > m_Circ(c2%).data(0).data0.radii < 3 Then
Else
End If
End Function
Public Function is_tangent_line_in_database(t_line As tangent_line_type, n%) As Boolean
For n% = 1 To last_conditions.last_cond(1).tangent_line_no
 If t_line.tangent_type = tangent_line(n%).tangent_type Then
    If t_line.tangent_type = tangent_line_by_point_on_circle And _
        read_point(t_line.data(0).coordinate(0), 0) = _
          tangent_line(n%).data(0).poi(0) Then
        is_tangent_line_in_database = True
         Exit Function
    End If
 ElseIf t_line.tangent_type = tangent_line_by_point_out_off_circle12 Or _
          t_line.tangent_type = tangent_line_by_point_out_off_circle21 Then
   If t_line.tangent_type = tangent_line(n%).tangent_type And _
       t_line.data(0).poi(0) = tangent_line(n%).data(0).poi(0) Then
        is_tangent_line_in_database = True
         Exit Function
   End If
  Else
   If t_line.tangent_type = tangent_line(n%).tangent_type And _
      t_line.data(0).ele(0).ty = tangent_line(n%).data(0).ele(0) And _
        t_line.data(0).ele(0).no = tangent_line(n%).data(0).ele(0).no And _
         t_line.data(0).ele(0).ty = tangent_line(n%).data(1).ele(0) And _
          t_line.data(0).ele(0).no = tangent_line(n%).data(1).ele(0).no Then
           is_tangent_line_in_database = True
            Exit Function
   End If
  End If
 End If
Next n%
End Function

Public Sub set_temp_circle_no(no%)
If draw_step < 3 Then
 temp_circle(0) = no%
Else
 temp_circle(1) = no%
End If
End Sub
Public Function redraw_point(ByVal point_no%, new_color As Byte, new_coord As POINTAPI)
Dim is_change As Boolean
If m_poi(point_no%).data(0).data0.color <> new_color Then
   m_poi(point_no%).data(0).data0.color = new_color
    is_change = True
End If
If (new_coord.X <> pointapi0.X Or new_coord.Y <> pointapi0.Y) And _
     (m_poi(point_no%).data(0).data0.coordinate.X <> new_coord.X Or _
       m_poi(point_no%).data(0).data0.coordinate.Y <> new_coord.Y) Then
   m_poi(point_no%).data(0).data0.coordinate = new_coord
    is_change = True
End If
If is_change Then
 Call C_display_picture.redraw_point(point_no%)
End If

   
End Function

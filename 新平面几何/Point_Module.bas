Attribute VB_Name = "Point_Module"
Option Explicit
Type point_data0_type
 name As String '* 1
  coordinate As POINTAPI
   color As Byte
    visible As Byte
End Type
Type point_data_type
 data0 As point_data0_type
  parent As parent_data_type
sons As son_data
 next_no As Integer '用于形成链表
  condition_ty As Byte
'    depend_element As depend_element_type
     depend_poi(8) As Integer
      from_wenti_no As Integer '导出的条件号
       inform As String
        is_change As Boolean
         other_no As Integer
      degree As Byte '自由度=0不受限（自由点）=1 半自由点（直线后圆上的自由点），=2自由点
  degree_for_reduce As Byte
in_line(26) As Integer '记录构图的过程
in_circle(7) As Integer
circle_center(8) As Integer '记录以此点为圆心的圆
g_line(1) As Integer '确定非自由点的直线
g_circle(1) As Integer
display As Boolean
End Type
Type point_type
is_set_data As Boolean
in_epolygon_no As Integer '
data(8) As point_data_type
End Type
Type io_point_data_type
point_data As point_data0_type
condition As Byte
point_no As Integer
depend_elemant As depend_element_type
End Type
Global m_poi() As point_type
Global m_aid_poi() As point_type
Public Sub PointCoordinateChange(point_no As Integer)

End Sub
Public Function m_point_number(p_coord As POINTAPI, ByVal ty As Byte, ByVal visible As Byte, ByVal color As Byte, _
                           name As String, d_ele1 As condition_type, d_ele2 As condition_type, _
                            inter_ty As Integer, is_set_data As Boolean, _
                              Optional old_or_new As Boolean = True) As Integer
Dim p_data0 As point_data_type
Dim p%, i%
m_point_number = read_point(p_coord, 0) '读点
If m_point_number > 0 Then '已有点
    old_or_new = False
    Exit Function
End If
'*******************************************************************
'建立新点
If name <> "" Then '由名获取点
  For i% = 1 To last_conditions.last_cond(1).point_no '***
   If m_poi(i).data(0).data0.name = name Or _
       is_same_POINTAPI(p_coord, m_poi(i%).data(0).data0.coordinate) Then
    m_point_number = i
     Exit Function
   End If
 Next i
End If
If is_set_data Then '设置点的数据
Call new_point_no(p%) '设置数据
p_data0.data0.color = color
p_data0.data0.visible = visible
p_data0.data0.coordinate = p_coord
p_data0.degree = 2
'If inter_ty > 0 Then
'p_data0.parent.inter_type = inter_ty
'End If
m_poi(p%).data(0) = p_data0
If name = "" Then
    m_poi(p%).data(0).data0.name = next_char(p%, "", 0, 0)
    name = m_poi(p%).data(0).data0.name
Else
    m_poi(p%).data(0).data0.name = name
End If
m_point_number = p%
         If d_ele1.ty = 0 And d_ele2.ty = 0 Then
            m_poi(p%).data(0).parent.inter_type = inter_ty
         Else
            If d_ele1.ty = line_ Then
              Call set_parent(d_ele1.ty, d_ele1.no, point_, p%, new_point_on_line)
            End If
            If d_ele2.ty = line_ Then
              Call set_parent(d_ele2.ty, d_ele2.no, point_, p%, new_point_on_line)
            End If
         End If
C_display_picture.draw_point (p%)
End If
End Function
Public Sub set_point_data0(point_no%, data0 As point_data0_type, is_change As Boolean, Optional degree As Byte = 2)
Call new_point_no(point_no%)
If data0.visible > 1 Then
   data0.visible = 1
End If
If data0.color = 0 Then
   data0.color = condition_color
End If
m_poi(point_no%).data(0).data0 = data0
m_poi(point_no%).data(0).is_change = is_change
m_poi(point_no%).data(0).degree = degree
m_poi(point_no%).data(0).degree_for_reduce = 0
If m_poi(point_no%).data(0).data0.color = 0 Then
   m_poi(point_no%).data(0).data0.color = 9
End If
Call C_display_picture.draw_point(point_no%)
'm_poi(point_no%).data(0).is_change = False
End Sub
Public Sub set_point_data_from_input(io_p_data As io_point_data_type)
Call new_point_no(io_p_data.point_no)
Call set_point_data0(io_p_data.point_no, io_p_data.point_data, True)
m_poi(io_p_data.point_no).data(0).depend_element = io_p_data.depend_elemant
m_poi(io_p_data.point_no).data(0).condition_ty = io_p_data.condition
End Sub
Public Sub set_point_data(point_no%, data0 As point_data_type, is_change As Boolean)
Call set_point_data0(point_no%, data0.data0, is_change)
m_poi(point_no%).data(0) = data0
End Sub
Public Sub set_point_color(point_no%, col As Byte)
Call new_point_no(point_no%)
If m_poi(point_no%).data(0).data0.color <> col Then
   m_poi(point_no%).data(0).data0.color = col
   Call C_display_picture.set_m_point_color(point_no%, col)
End If
End Sub
Public Sub new_point_no(point_no%) '申请新点的序号
Dim i%
If point_no > 0 Then '
    If last_conditions.last_cond(1).point_no < point_no Then '指定点超出已知点数
     ReDim Preserve m_poi(point_no Mod 10 + 10) As point_type
      last_conditions.last_cond(1).point_no = point_no
    End If
ElseIf point_no% = 0 Then '设新点
   For i% = 1 To last_conditions.last_cond(1).point_no
       If m_poi(i%).is_set_data = False Then '已知的范围有空号
         point_no% = i%
          Exit Sub
       End If
   Next i%
   If last_conditions.last_cond(1).point_no Mod 100 = 0 Then '没有空号设新点
    ReDim Preserve m_poi(last_conditions.last_cond(1).point_no + 100) As point_type
  End If
  last_conditions.last_cond(1).point_no = last_conditions.last_cond(1).point_no + 1
   point_no = last_conditions.last_cond(1).point_no
End If
   m_poi(point_no%).is_set_data = True
End Sub
Function read_point(in_coord As POINTAPI, ty As Byte) As Integer 'ty=0 down or up ,ty=1 move
Dim m%, i%, j%, l%, k%
Dim tp(3) '　判断鼠标点附近的点
'On Error GoTo read_point_error
read_point = 0
For k% = 1 To last_conditions.last_cond(1).point_no
If is_same_point_by_coord(m_poi(k%).data(0).data0.coordinate, in_coord) Then
read_point = k%
  If m_poi(k%).data(0).inform = "" Then
  Draw_form.Label2.Caption = LoadResString_(1255, "")
  Else
  Draw_form.Label2.Caption = m_poi(k%).data(0).inform
  End If
Exit Function
End If
Next k%
If ty = 1 Then
For k% = 1 To last_aid_point
  If Abs(m_poi(aid_point(k%)).data(0).data0.coordinate.X - in_coord.X) < 9 And _
               Abs(m_poi(aid_point(k%)).data(0).data0.coordinate.Y - in_coord.Y) < 9 Then
                  If m_poi(aid_point(k%)).data(0).other_no > 0 Then
                     read_point = m_poi(aid_point(k%)).data(0).other_no
                  'Else
                  '    m_point_data0 = m_poi(aid_point(k%)).data(0).data0
                  '    m_point_data0.color = 9
                  '    read_point = 0
                  '    Call set_point_data0(read_point, m_point_data0, True)
                  '    Call get_new_char(read_point)
                  '     m_poi(aid_point(k%)).data(0).other_no = read_point
                  End If
                 For i% = 0 To 15
                   If temp_point(i%).no = aid_point(k%) Then
                       temp_point(i%).no = read_point
                         Exit Function
                   End If
                 Next i%
                  Exit Function
  End If
Next k%
End If
read_point_error:
End Function
Public Function add_point_to_line(ByVal p%, ByVal lin%, n%, dis As Boolean, _
          is_draw As Boolean, red_point_start%, record As record_data_type, _
           Optional is_input As Boolean = True, Optional is_need_re_draw As Boolean = True) As Byte
Dim i%, j%, t_p1%, tp%, tn%
Dim r!
Dim k%
Dim t_l_data As line_data_type
Dim t_A%
If lin% > last_conditions.last_cond(1).line_no Or p% = 0 Or lin% = 0 Then
   Exit Function
ElseIf is_point_in_points(p%, m_lin(lin%).data(0).data0.in_point) Then
   Exit Function
End If
tp% = p%
'p_visible = m_poi(Abs(tp%)).data(0).data0.visible
t_l_data = m_lin(lin%).data(0)
'检查tp%是否在线上
  If add_point_to_line_data(p%, t_l_data, n%, dis, is_draw, red_point_start%, t_l_data, is_need_re_draw) = False Then '新点插入到直线数据中，并重新设置直线的端点坐标
    Exit Function
  End If
 If t_l_data.data0.in_point(0) <> m_lin(lin%).data(0).data0.in_point(0) Then '
    m_lin(lin%).data(0) = t_l_data
     For i% = 1 To m_lin(lin%).data(0).data0.in_point(0)
          If m_lin(lin%).data(0).data0.in_point(i%) <> p% Then
           '设两点线
               Call set_line_from_two_point(tp%, t_l_data.data0.in_point(i%), 0, 0, lin%, 0, "", record)
          End If
    Next i%
    For i% = 1 To m_lin(lin%).data(0).data0.in_point(0) - 1
           For j% = i% + 1 To m_lin(lin%).data(0).data0.in_point(0)
            If m_lin(lin%).data(0).data0.in_point(j%) <> p% And m_lin(lin%).data(0).data0.in_point(i%) <> p% Then
            '设三点线
             add_point_to_line = set_three_point_on_line(t_l_data.data0.in_point(i%), _
                t_l_data.data0.in_point(j%), tp%, record, 0, 0, True)
                If add_point_to_line > 1 Then
                 Exit Function
                End If
             End If
           Next j%
    Next i%
 Else 'If m_lin(lin%).data(0).data0.in_point(0) = t_l.data0.in_point(0) Then
   Exit Function
 End If
   m_lin(lin%).data(0) = t_l_data
   If m_lin(lin%).data(0).data0.in_point(0) = 2 Then  '直线只有两点，可作为直线的生成点
      m_lin(lin%).data(0).data0.poi(0) = m_lin(lin%).data(0).data0.in_point(1)
      m_lin(lin%).data(0).data0.poi(1) = m_lin(lin%).data(0).data0.in_point(2)
      m_lin(lin%).data(0).data0.end_point_coord(0) = m_poi(m_lin(lin%).data(0).data0.poi(0)).data(0).data0.coordinate
      m_lin(lin%).data(0).data0.end_point_coord(1) = m_poi(m_lin(lin%).data(0).data0.poi(1)).data(0).data0.coordinate
   End If
   If is_input Then
    Call set_parent(line_, lin%, point_, p%, new_point_on_line)
   End If
   'If m_poi(p%).data(0).parent.co_degree < 2 Then
   'm_poi(p%).data(0).parent.co_degree = m_poi(p%).data(0).parent.co_degree + 1
   'm_poi(p%).data(0).parent.element(m_poi(p%).data(0).parent.co_degree).ty = line_
   'm_poi(p%).data(0).parent.element(m_poi(p%).data(0).parent.co_degree).no = lin%
   'End If
'******************************************************************************
   'If m_poi(p%).data(0).degree > 0 Then
   '  m_poi(p%).data(0).degree = m_poi(p%).data(0).data0 - 1
   'End If
'***********************************************************************************
   'If m_poi(p%).data(0).parent.co_degree = 1 Then
   '  If Abs(m_poi(m_lin(lin%).data(0).data0.poi(0)).data(0).data0.coordinate.X - _
   '         m_poi(m_lin(lin%).data(0).data0.poi(1)).data(0).data0.coordinate.X) > 5 Then
   '        m_poi(p%).data(0).parent.ratio = _
   '         (m_poi(p%).data(0).data0.coordinate.X - _
   ''           m_poi(m_poi(m_lin(lin%).data(0).data0.poi(1)).data(0).data0.coordinate.X)) / _
   '         (m_poi(m_lin(lin%).data(0).data0.poi(0)).data(0).data0.coordinate.X - _
   '            m_poi(m_lin(lin%).data(0).data0.poi(1)).data(0).data0.coordinate.X)
   '  Else
   '         m_poi(p%).data(0).parent.ratio = _
   '         (m_poi(p%).data(0).data0.coordinate.Y - _
   '           m_poi(m_poi(m_lin(lin%).data(0).data0.poi(1)).data(0).data0.coordinate.Y)) / _
   '         (m_poi(m_lin(lin%).data(0).data0.poi(0)).data(0).data0.coordinate.Y - _
   '            m_poi(m_lin(lin%).data(0).data0.poi(1)).data(0).data0.coordinate.Y)
   '  End If
   'Else
   '  m_poi(p%).data(0).ratio = 0
   'End If
   'Call C_display_picture.re_draw_line(lin%)
'***************************************************************
   If is_point_in_points(lin%, m_poi(tp%).data(0).in_line) = 0 Then
       m_poi(tp%).data(0).in_line(0) = m_poi(tp%).data(0).in_line(0) + 1
     m_poi(tp%).data(0).in_line(m_poi(tp%).data(0).in_line(0)) = lin%
   End If
     Call simple_data_for_add_point_to_line(p%, lin%)
             Call C_display_picture.re_draw_line(lin%)
End Function
Public Function add_point_to_line_data(ByVal p%, l_data As line_data_type, n%, dis As Boolean, is_draw As Boolean, _
                red_point_start%, out_line_data As line_data_type, Optional is_need_re_draw As Boolean = True) As Boolean '=f 无添加 =t添加新点
Dim i%, j%, tp%, tn%
Dim r!
Dim k%
Dim t_l As line_data_type
tp% = Abs(p%)
'p_visible = m_poi(Abs(tp%)).data(0).data0.visible
If p% = 0 Then
 Exit Function
 End If
t_l = l_data
'检查tp%是否在线上
n% = is_point_in_points(p%, t_l.data0.in_point)
 If n% > 0 Then '在线上
  out_line_data = t_l
   Exit Function 'add_point_to_line_data=false
 End If
For i% = 1 To t_l.data0.in_point(0) '增加的点不在此直线上，搜寻全线
     If compare_two_point(m_poi(tp%).data(0).data0.coordinate, _
          m_poi(t_l.data0.in_point(i%)).data(0).data0.coordinate, t_l.data0.poi(0), t_l.data0.poi(1), 0) = 1 Then
         '　找出第一个比tp%顺序大的点
         For j% = t_l.data0.in_point(0) To i% Step -1          ' 后移
          t_l.data0.in_point(j% + 1) = t_l.data0.in_point(j%)
          t_l.data0.color(j% + 1) = t_l.data0.color(j%)
         Next j%
              t_l.data0.in_point(i%) = p% '插入
                   If i% = 1 Then
                    If is_need_re_draw Then
                     t_l.data0.color(1) = t_l.data0.color(2) '新增第一段颜色与原来第一段同
                      t_l.is_change = 255 '直线延伸需要重画
                    Else
                     t_l.data0.color(1) = 0
                    End If
                   ElseIf (i% = 2 And t_l.data0.color(1) = 0) Or _
                            (i% = t_l.data0.in_point(0) - 1 And t_l.data0.color(t_l.data0.in_point(0) - 1) = 0) _
                              And is_need_re_draw = False Then
                               t_l.data0.color(i%) = 0
                   Else
                    If is_need_re_draw Then
                      t_l.data0.color(i%) = t_l.data0.color(i% - 1)
                    Else
                                           t_l.data0.color(i%) = 0
                    End If
                   End If
                    t_l.data0.in_point(0) = t_l.data0.in_point(0) + 1
                    If t_l.data0.in_point(0) = 2 Then
                       t_l.data0.color(2) = 0 '单点直线，前加一点，原来的点成为末点
                    End If
           n% = i%
           GoTo add_point_to_line_out
      End If
 Next i%
         t_l.data0.in_point(0) = t_l.data0.in_point(0) + 1
            t_l.data0.in_point(t_l.data0.in_point(0)) = p%
            If is_need_re_draw Then
             If t_l.data0.in_point(0) > 2 Then
               t_l.data0.color(t_l.data0.in_point(0) - 1) = _
                         t_l.data0.color(t_l.data0.in_point(0) - 2)  '最后一段的颜色，与原来最后一段同
             End If
             t_l.is_change = 255
            End If
             t_l.data0.color(t_l.data0.in_point(0)) = 0
            n% = t_l.data0.in_point(0)
add_point_to_line_out:
If t_l.is_change = 255 And t_l.data0.in_point(0) >= 2 Then 'And t_l.data0.poi(0) = 0 Then '设置端点
      t_l.data0.poi(0) = t_l.data0.in_point(1)
       t_l.data0.poi(1) = t_l.data0.in_point(t_l.data0.in_point(0))
End If
'***********************************************************************
j% = 0
For i% = 1 To t_l.data0.in_point(0)
  If t_l.data0.in_point(i%) < 90 Then
   j% = j% + 1
    If j% >= 2 Then
    GoTo add_point_to_line_out3
    End If
  End If
Next i%
GoTo add_point_to_line_data_out
add_point_to_line_out3:
i% = 1
Do
If t_l.data0.in_point(i%) > 90 Then
   t_l.data0.in_point(0) = t_l.data0.in_point(0) - 1
   For j% = i% To t_l.data0.in_point(0)
    t_l.data0.in_point(j%) = t_l.data0.in_point(j% + 1)
   Next j%
  ' t_l.in_point(j%) = 0
   t_l.data0.in_point(j%) = 0
Else
i% = i% + 1
End If
Loop Until i% > t_l.data0.in_point(0)
add_point_to_line_data_out:
If t_l.data0.visible = 0 And dis Then
   t_l.data0.visible = 1
End If
'If t_l.data0.type = aid_condition Then
 '  t_l.data0.type = condition
  ' t_l.data0.color = condition_color
'End If
out_line_data = t_l
add_point_to_line_data = True '增加了新点
add_point_to_line_out1:
End Function

Public Function set_parent(ByVal ty As Integer, ByVal no%, son_ty As Integer, son_no As Integer, inter_point_type As Integer, _
                                Optional related_p1 As Integer = 0, Optional related_p2 As Integer = 0, _
                                 Optional related_point_start As Integer = 0) As Boolean
Dim i%, j%, t%
Dim t_coord As POINTAPI
Dim in_parent_data As parent_data_type
Dim son_sons As son_data
If ty = 0 Or no% = 0 Or son_ty = 0 Or son_no = 0 Then
   Exit Function
End If
'*******************************************************************
If son_ty = point_ Then '设置子数据
 in_parent_data = m_poi(son_no).data(0).parent
 If in_parent_data.co_degree >= 2 Then
  Exit Function
 End If
 son_sons = m_poi(son_no).data(0).sons
ElseIf son_ty = line_ Then
 in_parent_data = m_lin(son_no).data(0).parent
 son_sons = m_lin(son_no).data(0).sons
ElseIf son_ty = circle_ Then
 in_parent_data = m_Circ(son_no).data(0).parent
 son_sons = m_Circ(son_no).data(0).sons
ElseIf son_ty = wenti_cond_ Then
 GoTo set_parent_mark1 '如果ty=wenti_cond_ 不能有parent
ElseIf son_ty = 0 Then
 Exit Function
End If
'在in_parat_data中添加(ty,no%),在son_sons中删除(ty,no%)
'******************************************************************************
'If ty = 0 Or no% = 0 Then    '已有两个确定点的条件
'   Exit Function
'Else '/1 if
 For i% = 1 To son_sons.last_son '搜索sons中数据
     If son_sons.son(i%).ty = ty And son_sons.son(i%).no = no% Then
      son_sons.last_son = son_sons.last_son - 1
       For j% = i% To son_sons.last_son
           son_sons.son(j%) = son_sons.son(j% + 1)
       Next j%
     End If
 Next i%
 For i% = 1 To in_parent_data.last_element '搜索parent中数据
  If in_parent_data.element(i%).ty = ty And _
      in_parent_data.element(i%).no = no% Then '已有数据
      If related_p1 > 0 Or related_p2 > 0 Then
        GoTo set_parent_mark2
      Else
       Exit Function '已此条件有
      End If
  End If
 Next i%
 '排序
    '******************************************************************************************************************
           in_parent_data.last_element = in_parent_data.last_element + 1 '增加新数据
           If in_parent_data.co_degree < 2 Then
              in_parent_data.co_degree = in_parent_data.co_degree + 1
           End If
      For i% = 1 To in_parent_data.last_element - 1 '排序
       If (ty < in_parent_data.element(i%).ty) Or (ty = in_parent_data.element(i%).ty And no% < in_parent_data.element(i%).no) Then
         For j% = in_parent_data.last_element To i% + 1 Step -1
          in_parent_data.element(j%) = in_parent_data.element(j% - 1)
         Next j%
          in_parent_data.element(i%).ty = ty
          in_parent_data.element(i%).no = no
        GoTo set_parent_data_out
       End If
      Next i%
      in_parent_data.element(in_parent_data.last_element).ty = ty
      in_parent_data.element(in_parent_data.last_element).no = no%
set_parent_data_out:
'*******************************************************************************
set_parent_mark2:
'*******************************************************************************
      If related_p1 > 0 Or related_p2 > 0 Then
         'if related_p1>rel
         If is_element_in_parent(point_, related_p1%, m_lin(no%).data(0).parent) = False Then
           Call set_parent(point_, related_p1%, son_ty, son_no, 0)
         End If
         If is_element_in_parent(point_, related_p2%, m_lin(no%).data(0).parent) = False Then
           Call set_parent(point_, related_p2%, son_ty, son_no, 0)
         End If
         in_parent_data.related_point(0) = related_p1
         in_parent_data.related_point(1) = related_p2
         in_parent_data.related_point(2) = related_point_start%
      ElseIf son_ty = point_ And ty = line_ And m_lin(in_parent_data.element(1).no).data(0).data0.depend_poi(0) > 0 _
         And m_lin(in_parent_data.element(1).no).data(0).data0.depend_poi(1) > 0 Then
         in_parent_data.related_point(0) = m_lin(in_parent_data.element(1).no).data(0).data0.depend_poi(0)
         in_parent_data.related_point(1) = m_lin(in_parent_data.element(1).no).data(0).data0.depend_poi(1)
         in_parent_data.related_point(2) = m_lin(in_parent_data.element(1).no).data(0).data0.depend_poi(0)
         'in_parent_data.inter_type = paral_
      End If
'***************************************************************************************************
         If inter_point_type <> 0 Then
                  in_parent_data.inter_type = new_inter_point_type(in_parent_data.inter_type, inter_point_type)
         End If
         'If (ty <> paral_ And ty <> verti_ And ty <> related_line_ And ty <> Ratio_for_measure_) And son_ty = point_ Then '非直接强制条件
         '  If in_parent_data.co_degree < 2 Then
         ' in_parent_data.co_degree = in_parent_data.co_degree + 1 '增加新的约束条件
         ' End If
         'End If
      If inter_point_type = length_depended_by_two_points_ And son_ty = circle_ Then
              m_Circ(son_no).data(0).data0.real_radii = _
               distance_of_two_POINTAPI(m_poi(related_p1).data(0).data0.coordinate, _
                 m_poi(related_p2).data(0).data0.coordinate)
      'ElseIf son_ty = point_ And ty = line_ And inter_point_type > 0 Then
      '   in_parent_data.co_degree = 2
      End If
      
'******************************************************************************
     If in_parent_data.element(1).ty = line_ Then  '若是第一约束（半约束），并是直线，记录点在直线上的相对位置
       If in_parent_data.element(2).no = 0 And (in_parent_data.inter_type = new_point_on_line Or _
            in_parent_data.inter_type = paral_ Or in_parent_data.inter_type = verti_ Or _
              in_parent_data.inter_type = new_point_on_line) Then
          in_parent_data.ratio = get_ratio_of_point_on_line(son_no, in_parent_data.element(1).no, _
                 in_parent_data.related_point(0), in_parent_data.related_point(1), in_parent_data.related_point(2))
       End If
      'If in_parent_data.related_point(0) > 0 And in_parent_data.related_point(1) > 0 Then
      '       t_coord = minus_POINTAPI(m_poi(in_parent_data.related_point(1)).data(0).data0.coordinate, _
                                m_poi(in_parent_data.related_point(0)).data(0).data0.coordinate)
      '      If in_parent_data.element(2).ty = verti_ Then
      '        t_coord = verti_POINTAPI(t_coord)
      '       End If
      '     If Abs(t_coord.X) > 5 Then
      '        in_parent_data.ratio = _
               (m_poi(son_no).data(0).data0.coordinate.X - _
                 m_poi(in_parent_data.related_point(0)).data(0).data0.coordinate.X) / t_coord.X
      '     Else
      '        in_parent_data.ratio = _
      '         (m_poi(son_no).data(0).data0.coordinate.Y - _
      '           m_poi(in_parent_data.related_point(0)).data(0).data0.coordinate.Y) / t_coord.Y
      '     End If
      'End If
   End If
 '  End If
'*******************************************************************************************************
'If son_ty = point_ Then
  'If in_parent_data.co_degree = 1 And inter_point_type > 0 Then
  '   in_parent_data.inter_type = new_inter_point_type(inter_point_type, in_parent_data.inter_type)
  'If in_parent_data.co_degree = 2 Then
  '   If in_parent_data.element(1).ty = line_ And in_parent_data.element(2).ty = line_ _
                                And in_parent_data.inter_type <> interset_point_line_line Then
  '      in_parent_data.inter_type = new_inter_point_type(in_parent_data.inter_type, interset_point_line_line)
  '   ElseIf in_parent_data.element(1).ty = line_ And in_parent_data.element(2).ty = circle_ And _
       (in_parent_data.inter_type <> new_point_on_line_circle12 Or in_parent_data.inter_type <> new_point_on_line_circle21) Then
  '     i% = inter_point_line_circle(in_parent_data.element(1).no, in_parent_data.element(2).no, _
                    m_poi(son_no).data(0).data0.coordinate, son_no, False, False)
  '                  If t% > 0 Then
  '                   in_parent_data.inter_type = new_inter_point_type(t%, in_parent_data.inter_type)
  '                  End If
  '   ElseIf in_parent_data.element(1).ty = circle_ And in_parent_data.element(2).ty = circle_ And _
       (in_parent_data.inter_type <> new_point_on_circle_circle12 And in_parent_data.inter_type <> new_point_on_circle_circle21 And _
         in_parent_data.inter_type <> tangent_point_) Then
  '     t% = inter_point_circle_circle(in_parent_data.element(1).no, in_parent_data.element(2).no, t_coord)
  '     If t% > 0 Then
  '        in_parent_data.inter_type = new_inter_point_type(t%, in_parent_data.inter_type)
  '     End If
  '   End If
  'End If
'End If
'**************************************************************************************************
set_parent_mark1:
If son_ty = point_ Then '设置子数据
 m_poi(son_no).data(0).parent = in_parent_data
 m_poi(son_no).data(0).sons = son_sons
ElseIf son_ty = line_ Then
 m_lin(son_no).data(0).parent = in_parent_data
 m_lin(son_no).data(0).sons = son_sons
ElseIf son_ty = circle_ Then
 m_Circ(son_no).data(0).parent = in_parent_data
 m_Circ(son_no).data(0).sons = son_sons
'ElseIf son_ty = wenti_cond_ Then
' GoTo set_parent_mark1 '如果ty=wenti_cond_ 不能有parent
End If
'*******************************************************************************************
'*********************************************************************************************
'设置son数据
If ty = point_ Then
   son_sons = m_poi(no%).data(0).sons
   in_parent_data = m_poi(no%).data(0).parent
ElseIf ty = line_ Then
   son_sons = m_lin(no%).data(0).sons
   in_parent_data = m_lin(no%).data(0).parent
ElseIf ty = circle_ Then
   son_sons = m_Circ(no%).data(0).sons
   in_parent_data = m_Circ(no%).data(0).parent
ElseIf ty = Ratio_for_measure_ Then
   son_sons = Ratio_for_measure.sons
Else
  Exit Function
End If
'在中son_sons中添加(son_ty,son_no)，在parent中删除(son_ty,son_no)
For i% = 1 To in_parent_data.last_element
    If in_parent_data.element(i%).ty = son_ty And in_parent_data.element(i%).no = son_no Then
       in_parent_data.last_element = in_parent_data.last_element - 1
        For j% = i% To in_parent_data.last_element
          in_parent_data.element(j%) = in_parent_data.element(j% + 1)
        Next j%
    End If
Next i%
For i% = 1 To son_sons.last_son '如果，son_sons含father,必修删除，否则会产生循环
  If son_sons.son(i%).ty = son_ty And son_sons.son(i%).no = son_no% Then
   Exit Function
  End If
Next i%
'***************************************************************
son_sons.last_son = son_sons.last_son + 1
      son_sons.son(son_sons.last_son).ty = son_ty
      son_sons.son(son_sons.last_son).no = son_no%
If ty = point_ Then '回传
   m_poi(no%).data(0).sons = son_sons
   m_poi(no%).data(0).parent = in_parent_data
ElseIf ty = line_ Then
   m_lin(no%).data(0).sons = son_sons
   m_lin(no%).data(0).parent = in_parent_data
ElseIf ty = circle_ Then
   m_Circ(no%).data(0).sons = son_sons
   m_Circ(no%).data(0).parent = in_parent_data
ElseIf ty = Ratio_for_measure_ Then
   Ratio_for_measure.sons = son_sons
End If
'End If
End Function

Private Sub add_sons_data(son_ty As Byte, son_no As Integer, sons_data As son_data)
Dim i%
For i% = 1 To sons_data.last_son
    If sons_data.son(i%).ty = son_ty And sons_data.son(i%).no = son_no Then
        Exit Sub
    End If
Next i%
sons_data.last_son = sons_data.last_son + 1
sons_data.son(sons_data.last_son).ty = son_ty
sons_data.son(sons_data.last_son).no = son_no

End Sub


Public Function is_element_in_parent(ele_ty As Byte, ele_no%, parent_data As parent_data_type) As Boolean
Dim i%
For i% = 1 To parent_data.last_element
 If parent_data.element(i%).ty = ele_ty And parent_data.element(i%).no = ele_no% Then
  is_element_in_parent = True
   Exit Function
 End If
Next i%
End Function

Public Function new_inter_point_type(inter_point_type1 As Integer, inter_point_type2 As Integer) As Integer
If inter_point_type2 = 0 Then
   new_inter_point_type = inter_point_type1
ElseIf inter_point_type1 = 0 Then
   new_inter_point_type = inter_point_type2
ElseIf inter_point_type1 = interset_point_line_line Or _
     inter_point_type1 = new_point_on_circle_circle12 Or _
         inter_point_type1 = new_point_on_circle_circle21 Or _
          inter_point_type1 = new_point_on_line_circle12 Or _
           inter_point_type1 = new_point_on_line_circle21 Or _
          inter_point_type1 = paral_ Or _
           inter_point_type1 = verti_ Or _
            inter_point_type1 = tangent_line_by_point_out_off_circle12 Or _
             inter_point_type1 = tangent_line_by_point_out_off_circle21 Then
   new_inter_point_type = inter_point_type1
ElseIf inter_point_type2 = interset_point_line_line Or _
    inter_point_type2 = new_point_on_circle_circle12 Or _
         inter_point_type2 = new_point_on_circle_circle21 Or _
          inter_point_type2 = new_point_on_line_circle12 Or _
           inter_point_type2 = new_point_on_line_circle21 Or _
          inter_point_type2 = paral_ Or _
            inter_point_type2 = verti_ Or _
             inter_point_type2 = tangent_line_by_point_out_off_circle12 Or _
              inter_point_type2 = tangent_line_by_point_out_off_circle21 Then
   new_inter_point_type = inter_point_type2
ElseIf inter_point_type2 = new_point_on_line And _
   inter_point_type2 = new_point_on_line Then
   new_inter_point_type = interset_point_line_line
Else
   new_inter_point_type = inter_point_type1
'ElseIf (inter_point_type1 = new_point_on_line Or inter_point_type1 = exist_point) And _
   (inter_point_type2 = new_point_on_line_circle12 Or _
     inter_point_type2 = new_point_on_line_circle21 Or _
      inter_point_type2 = new_point_on_line) Then
'   new_point_type = inter_point_type2
'ElseIf (inter_point_type2 = new_point_on_line Or inter_point_type2 = exist_point) And _
   (inter_point_type1 = new_point_on_line_circle12 Or _
     inter_point_type1 = new_point_on_line_circle21 Or _
      inter_point_type1 = new_point_on_line) Then
'   new_point_type = inter_point_type1
End If
End Function

Attribute VB_Name = "line_Module"
Option Explicit
Public Type display_line_data
end_point_coord(1) As POINTAPI
end_point(1) As Integer
color As Byte
display As Byte '0,1画线,2，消线
End Type
Sub set_line_data_from_input(io_line_data As io_line_data_type)
Dim temp_line As line_data_type
Dim i%
temp_line.data0 = io_line_data.line_data
temp_line.condition = io_line_data.condition
temp_line.depend_element = io_line_data.depend_element
For i% = 0 To io_line_data.line_data.in_point(0)
temp_line.in_point(i%) = Abs(io_line_data.line_data.in_point(i%))
Next i%
Call set_line_data0(io_line_data.line_no, temp_line, 0, 0)
End Sub
Public Function line_number0(ByVal p1%, ByVal p2%, n1%, n2%, Optional is_just_jug As Boolean = True) As Integer
'供推理用 is_reduce=false 生成新的直线
Dim i%, k%, tn%
Dim temp_record As record_data_type
If p1% = 0 Or p2% = 0 Or p1% = p2% Then
 line_number0 = 0
  Exit Function
Else
  line_number0 = search_for_line_number_from_two_point(p1%, p2%, n1%, n2%) '搜索过的直线
If line_number0 = 0 And is_just_jug = False Then '无过p1,p2两点的直线
'建立新直线
   line_number0 = set_line(p1%, p2%, n1%, n2%, m_poi(p1%).data(0).data0.coordinate, _
                                       m_poi(p2%).data(0).data0.coordinate, pointapi0, _
                                        depend_condition(point_, p1%), _
                                          depend_condition(point_, p2%), _
                           condition, condition_color, 0, 0)
    If line_number0 > 0 Then '添加两点搜索的直线数据库
     Call set_line_from_two_point(p1%, p2%, n1%, n2%, line_number0, 0, "", temp_record)
    End If
End If
End If
End Function
Private Function is_different_line_data0(line_data1 As line_data0_type, line_data2 As line_data0_type) As Boolean
Dim i%
 If line_data1.total_color <> line_data2.total_color Or line_data1.in_point(0) <> line_data2.in_point(0) Then
     is_different_line_data0 = True
 ElseIf is_same_two_point(line_data1.poi(0), line_data1.poi(1), line_data2.poi(0), line_data2.poi(1)) = False Then
     is_different_line_data0 = True
 ElseIf line_data1.visible <> line_data2.visible Then
     is_different_line_data0 = True
 ElseIf line_data1.type <> line_data2.type Then
     is_different_line_data0 = True
 Else
     For i% = 1 To line_data1.in_point(0)
      If line_data1.in_point(i%) <> line_data2.in_point(i%) Then
        is_different_line_data0 = True
         Exit Function
      End If
     Next i%
 End If
End Function
Function line_number(ByVal p1%, p2%, p_coord1 As POINTAPI, p_coord2 As POINTAPI, _
                            d_ele1 As condition_type, d_ele2 As condition_type, _
                             line_type As Byte, color As Byte, visible As Byte, _
                              dir As Integer, Optional draw_type As Byte = 0) As Integer
'p1%,p2% 确定直线的两点的序号,draw_type=1 画线=2画线但不重画
'p_cood1,p_cood2 两点的坐标
'line_type 直线的类型,condition 或 conclusion
'ty1 记录条件，ty2　是否显示
Dim i%, n1%, n2%, l%
If p1% = p2% And p1% > 0 Then '两个相同实点,不能确定直线退出
  Exit Function
Else
     l% = search_for_line_number_from_two_point(p1%, p2%, n1%, n2%) '由两点确定已有直线或建立新直线
      If l% > 0 Then 'ty1 = condition Then
        line_number = l% '搜索到过p1%,p2%的直线
       If m_lin(line_number).data(0).data0.visible <> visible Then '如果原线不可示,根据visible 重新设置
          m_lin(line_number).data(0).data0.visible = visible
           m_lin(line_number).data(0).data0.type = line_type
            m_lin(line_number).data(0).is_change = 255
       End If
       If p2% > 0 Then
        If n1% > n2% Then
            Call exchange_two_integer(n1%, n2%)
        End If
        If m_lin(line_number).data(0).data0.type = condition And line_type = conclusion Then '原来是condition现在是conclusion重新设置
          m_lin(line_number).data(0).data0.type = conclusion
           For i% = n1% To n2% - 1 '将结论部分变为红色,负点号开始的线段为红色
            m_lin(line_number).data(0).data0.color(i%) = conclusion_color
            'If m_lin(line_number).data(0).data0.in_point(i%) > 0 Then '
            ' m_lin(line_number).data(0).data0.in_point(i%) = -m_lin(line_number).data(0).data0.in_point(i%)
            'End If
          Next i%
          m_lin(line_number).data(0).data0.total_color = conclusion_color
          m_lin(line_number).data(0).is_change = 255
set_color_out:
        ElseIf m_lin(line_number).data(0).data0.type = condition Then
           m_lin(line_number).data(0).data0.total_color = condition_color
        ElseIf m_lin(line_number).data(0).data0.type = conclusion And draw_type = 1 Then
              m_lin(line_number).data(0).data0.type = condition
           For i% = n1% To n2%
            m_lin(line_number).data(0).data0.color(i%) = conclusion_color
            'If m_lin(line_number).data(0).data0.in_point(i%) < 0 Then
            '  m_lin(line_number).data(0).data0.in_point(i%) = -m_lin(line_number).data(0).data0.in_point(i%)
            '   m_lin(line_number).data(0).data0.total_color = condition_color
            'End If
           Next i%
           m_lin(line_number).data(0).data0.total_color = conclusion_color
          End If
           Call C_display_picture.re_draw_line(line_number)
        Else 'p2%<=0
           Call C_display_picture.re_draw_line(line_number)
       End If
     Else 'l%
      line_number = set_line(p1%, p2%, n1%, n2%, p_coord1, p_coord2, p_coord2, d_ele1, d_ele2, line_type, color, _
              visible, dir, draw_type)  '
          '设置直线的前后辈
     End If
  End If
'End If
 End Function

Public Function set_line(ByVal p1%, ByVal p2%, n1%, n2%, coord1 As POINTAPI, coord2 As POINTAPI, _
                         p_coord As POINTAPI, _
                         depend_cond1 As condition_type, _
                         depend_cond2 As condition_type, _
                         line_ty As Byte, color As Byte, visible As Byte, _
                                         dir As Integer, Optional draw_type As Byte) As Integer
Dim i%, j%, k%, temp_no%, bra%
Dim temp_coord(1) As POINTAPI
Dim tl(1) As Integer
Dim n_(1) As Integer
Dim temp_record As record_data_type
Dim c_data0 As condition_data_type
Dim line_data_0 As line_data_type
If p1% = p2% Or p1% <= 0 Then '两个端点同，起点的序号为负
 Exit Function
ElseIf line_number0(p1%, p2%, 0, 0) > 0 Then
 Exit Function
'*******************************
ElseIf p1% > 0 And p2% > 0 Then '两点序号为正
      set_line = line_number0(p1%, p2%, n1%, n2%)
       If set_line > 0 Then
       If m_lin(set_line).data(0).data0.visible = 0 And visible = 1 Then
          m_lin(set_line).data(0).data0.visible = 1
           m_lin(set_line).data(0).is_change = 255
       Else
          If n1% > n2% Then
           Call exchange_two_integer(n1%, n2%)
          End If
          m_lin(set_line).data(0).data0.color(n1%) = color
          m_lin(set_line).data(0).is_change = 255
       End If
       If m_lin(set_line).data(0).data0.type <> line_ty Then
          m_lin(set_line).data(0).data0.type = line_ty
          m_lin(set_line).data(0).is_change = 255
       End If
       n_(0) = n1%
       n_(1) = n2%
       If n_(0) > n_(1) Then
         Call exchange_two_integer(n_(0), n_(1))
       End If
       For k% = n_(0) To n_(1) - 1
         If m_lin(set_line).data(0).data0.color(k%) <> color Then
            m_lin(set_line).data(0).data0.color(k%) = color
                   m_lin(set_line).data(0).is_change = 255
         End If
       Next k%
        If m_lin(set_line).data(0).is_change = 255 Then
            Call C_display_picture.re_draw_line(set_line)
        End If
       Exit Function '退出，输出直线序号
     End If
  '  Next j%
  'Next i%
  '************************************************************
  '此两点是生成直线的基本点，先排序，以后生成的点都以它们为标准排序
   k% = compare_two_point(m_poi(p1%).data(0).data0.coordinate, _
       m_poi(p2%).data(0).data0.coordinate, 0, 0, 6) '按规定规则排序
   If k% = 0 Then
     Exit Function '无法排序，退出
   ElseIf k% = -1 Then
    Call exchange_two_integer(p1%, p2%) '逆序，交换点
     dir = dir * -1 '设置方向
   End If
'set_line = last_conditions.last_cond(1).line_no
  line_data_0.data0.in_point(0) = 2 '
'*************************************************************************************
   line_data_0.data0.in_point(1) = p1%
   line_data_0.data0.in_point(2) = p2%
    '设置原始生成点，除非他们本身发生变化，其他变化不会影响他们
   line_data_0.data0.end_point_coord(0) = m_poi(p1%).data(0).data0.coordinate
   line_data_0.data0.end_point_coord(1) = m_poi(p2%).data(0).data0.coordinate
   line_data_0.data0.poi(0) = p1%
   line_data_0.data0.poi(1) = p2%
   line_data_0.data0.depend_poi(0) = p1%
   line_data_0.data0.depend_poi(1) = p2%
   line_data_0.data0.depend_poi1_coord.X = -10000
   line_data_0.data0.depend_poi1_coord.Y = -10000
ElseIf p1% > 0 Then '第一点序号大于零
 If is_same_point_by_coord(m_poi(p1%).data(0).data0.coordinate, p_coord) Then
     Exit Function
 Else
  line_data_0.data0.in_point(0) = 1
  line_data_0.data0.in_point(1) = p1%
  line_data_0.data0.poi(0) = p1%
  line_data_0.data0.depend_poi(0) = p1%
  'line_data_0.data0.end_point_coord(0) = m_poi(p1%).data(0).data0.coordinate
  line_data_0.data0.depend_poi1_coord = coord2
  'line_data_0.data0.end_point_coord(1) = coord2
End If
End If
'***********************************************************
'设置画直线的端点坐标
'*********************************************************************************
line_data_0.data0.in_point(10) = dir
line_data_0.data0.color(1) = color
If color <> condition_color Then
   line_data_0.is_change = True '为了结束画图后，恢复
End If
line_data_0.data0.type = line_ty
line_data_0.data0.total_color = 0
'If line_ty = conclusion Then
'line_data_0.data0.in_point(1) = -line_data_0.data0.in_point(1)
'End If
line_data_0.data0.visible = visible
'*****************************************************************************
Call set_line_data0(temp_no%, line_data_0, p_coord.X, p_coord.Y, draw_type)
 'If depend_cond1.no > 0 Then
 '  Call add_d_condition_to_line(m_lin(temp_no%).data(0), depend_cond1)
   Call set_parent(depend_cond1.ty, depend_cond1.no, line_, temp_no, 0)
 'End If
' If depend_cond2.no > 0 Then
'   Call add_d_condition_to_line(m_lin(temp_no%).data(0), depend_cond2)
   Call set_parent(depend_cond2.ty, depend_cond2.no, line_, temp_no, 0)
' End If
Call set_line_direction(temp_no%, dir)
set_line = temp_no%
If k% = 1 Then
n1% = 1
 n2% = 2
ElseIf k% = -1 Then
n1% = 2
 n2% = 1
End If
Call set_point_in_line(p1%, temp_no%)
Call set_point_in_line(p2%, temp_no%)
'If run_statue > 1 Then '12.10
'   If m_lin(temp_no%).data(0).branch = 0 Then
'      Call set_element_branch(line_, temp_no%, m_poi(p1%).data(0).branch)
'   Else
'      If m_lin(temp_no%).data(0).branch > m_poi(p1%).data(0).branch Then
'         Call connect_two_branch(m_lin(temp_no%).data(0).branch, m_poi(p1%).data(0).branch)
'      ElseIf m_lin(temp_no%).data(0).branch < m_poi(p1%).data(0).branch Then
'         Call connect_two_branch(m_poi(p1%).data(0).branch, m_lin(temp_no%).data(0).branch)
'      End If
'   End If
'End If
For i% = 1 To m_lin(temp_no%).data(0).data0.in_point(0) - 1
 For j% = i% + 1 To m_lin(temp_no%).data(0).data0.in_point(0)
  Call set_line_from_two_point(m_lin(temp_no%).data(0).data0.in_point(i%), _
         m_lin(temp_no%).data(0).data0.in_point(j%), i%, j%, _
           temp_no%, 0, 0, temp_record)
Next j%
Next i%
'******************************************************************
Exit Function
set_line_error:
End Function
Public Sub set_line_data0(line_no%, Ldata0 As line_data_type, X As Long, Y As Long, Optional draw_type As Byte)
Dim t_data0 As line_data_type
Dim i%, j%, empty_no%
Dim temp_record As record_data_type
t_data0 = Ldata0 '输入新的直线数据
   If t_data0.data0.in_point(0) = 0 Then
     Exit Sub
    End If
 '  If last_conditions.last_cond(1).line_no Mod 100 = 0 And line_no% > last_conditions.last_cond(1).line_no Then '增加数据库的直线数据容量_
 '     ReDim Preserve m_lin(last_conditions.last_cond(1).line_no + 100) As line_type
 '    last_conditions.last_cond(1).line_no = line_no%
 '  End If
'If line_no% = 0 Then '输入的直线序号=0，新建直线，将新直线放入数据库，如发现已有相同的直线推出
' For i% = 1 To last_conditions.last_cond(1).line_no
'     If m_lin(i%).data(0).data0.in_point(0) > 0 Then
'      If is_two_line_same0(m_lin(i%).data(0).data0, t_data0.data0) > 0 Then
'        line_no = i%
'         For j% = 1 To Ldata0.data0.in_point(0)
'          Call add_point_to_line(Ldata0.data0.in_point(j%), line_no, 0, True, False, 0, temp_record)
'         Next j%
'         t_data0 = combine_two_line0(m_lin(line_no%).data(0), Ldata0)
'           m_lin(line_no%).data(0) = t_data0
'          If t_data0.is_change Then
'           'm_lin(line_no%).data(0).is_change = True
'           GoTo set_line_data0_out
'          End If
'          Exit Sub
'      End If
'     End If
 'Next i%
 'End If
'******************************************************************************
'未发现已有相同的直线
 If line_no% = 0 Then
  If last_conditions.last_cond(1).line_no = last_conditions.last_cond(2).line_no Then
   ReDim Preserve m_lin(last_conditions.last_cond(2).line_no + 100) As line_type
   last_conditions.last_cond(2).line_no = last_conditions.last_cond(2).line_no + 100
  End If
   last_conditions.last_cond(1).line_no = last_conditions.last_cond(1).line_no + 1 '新点
     line_no% = last_conditions.last_cond(1).line_no
  End If
 '***********************************************************************************
m_lin(line_no%).data(0) = t_data0
'm_lin(line_no%).data(0).is_change = True
m_lin(line_no%).data(0).other_no = line_no%
For i% = 1 To m_lin(line_no%).data(0).data0.in_point(0) - 1
 For j% = i% + 1 To m_lin(line_no%).data(0).data0.in_point(0)
  Call set_line_from_two_point(m_lin(line_no%).data(0).data0.in_point(i%), _
         m_lin(line_no%).data(0).data0.in_point(j%), i%, j%, _
           line_no%, 0, 0, temp_record)
Next j%
Next i%
set_line_data0_out:
If m_lin(line_no%).data(0).data0.in_point(0) > 2 Then
For i% = 1 To m_lin(line_no%).data(0).data0.in_point(0)
 Call add_line_to_point(line_no%, m_lin(line_no%).data(0).data0.in_point(i%))
Next i%
End If
'               Call draw_temp_line_for_input(1)
If m_lin(line_no%).data(0).data0.in_point(0) >= 1 Then
  m_lin(line_no%).data(0).is_change = 255
    Call C_display_picture.set_m_line_data0(line_no%, X, Y)
End If
End Sub
Public Sub add_d_condition_to_line(line_d As line_data_type, _
                                  d_cond1 As condition_type)
Dim com_ty As Integer
If d_cond1.no > 0 Then
   If line_d.parent.element(0).ty = line_ And _
        line_d.parent.element(0).ty = d_cond1.no Then
         Exit Sub
   ElseIf line_d.parent.element(1).ty = line_ And _
           line_d.parent.element(1).ty = d_cond1.no Then
         Exit Sub
   Else
    If line_d.parent.element(1).no > 0 Then
       Exit Sub
    Else 'If line_d.parent.element(1).no = 0 And _
            d_cond2.no = 0 Then
      com_ty = compare_two_conditions(d_cond1, line_d.parent.element(0))
      If com_ty = 1 Then
         line_d.parent.element(1) = line_d.parent.element(0)
         line_d.parent.element(0) = d_cond1
      ElseIf com_ty = -1 Then
         line_d.parent.element(1) = d_cond1
      End If
    End If
   End If
 End If
End Sub
Public Function compare_two_conditions(d_cond1 As condition_type, _
                                       d_cond2 As condition_type) As Integer
   If d_cond1.ty = 0 And d_cond2.ty = 0 Then
      Exit Function
   ElseIf (d_cond1.ty > d_cond2.ty And d_cond2.ty > 0) Or _
        (d_cond1.ty = 0 And d_cond2.ty > 0) Then
      compare_two_conditions = -1 '
   ElseIf (d_cond1.ty < d_cond2.ty And d_cond1.ty > 0) Or _
        (d_cond2.ty = 0 And d_cond1.ty > 0) Then
      compare_two_conditions = 1 '
   ElseIf d_cond1.ty = d_cond2.ty Then
     If d_cond1.no > d_cond2.no Then
        compare_two_conditions = -1
     ElseIf d_cond1.no < d_cond2.no Then
        compare_two_conditions = 1
     Else
        compare_two_conditions = 0
     End If
   End If
 End Function
Public Sub set_line_from_aid_line(ByVal in_l1%, p%, p_coord As POINTAPI, out_l%)
Dim tp%
Dim t_cond(1) As condition_type
p% = m_point_number(p_coord, condition, 1, condition_color, "", _
                    depend_condition(0, 0), depend_condition(0, 0), 0, False)
Call C_display_picture.get_line_depend_data(in_l1%, _
                          t_cond(0).ty, t_cond(0).no, t_cond(1).ty, t_cond(1).no)
Call C_display_picture.get_end_point(in_l1%, tp%, 0, 0, 0, 0, 0)
out_l% = line_number(tp%, p%, pointapi0, pointapi0, t_cond(0), t_cond(1), _
                   condition, condition_color, 1, 0)

End Sub

Public Function second_end_point_coordinate(line_no%) As POINTAPI
If m_lin(line_no%).data(0).data0.depend_poi(1) > 0 Then
  second_end_point_coordinate = _
      m_poi(m_lin(line_no%).data(0).data0.depend_poi(1)).data(0).data0.coordinate
Else
  second_end_point_coordinate = m_lin(line_no%).data(0).data0.depend_poi1_coord
End If
End Function

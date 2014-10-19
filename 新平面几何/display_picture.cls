VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "display_picture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Last_old_m_circle_count
Public m_point As New Collection
Public m_line As New Collection
Public m_circle As New Collection
Public m_conclusion_line As New Collection
Public m_conclusion_circle As New Collection
Public m_aid_line As New Collection
Public m_draw_statue As Byte
Private m_last_aid_line As Integer
Private draw_aid_line_poi As POINTAPI
Private m_object As Object
Private picture_max_range As Long
Private m_picture_top_left_x As Long
Private m_picture_top_left_y As Long
Private m_picture_bottom_right_x As Long
Private m_picture_bottom_right_y As Long
Private picture_width_height_x As Long
Private picture_width_height_y As Long
Private picture_range As Long
Private red_or_black As Boolean
Private is_set_print_data As Boolean
Private me_class As Object
Private last_draw_point As Integer
Private m_circ_data As circle_data_type
Private aid_line(1 To 4) As aid_line_data
Private last_aid_line As in_paral_type
Public Function read_aid_line0(ByVal in_x As Long, ByVal in_y As Long, out_x As Long, out_y As Long, _
                                  Optional ty As Byte = 0) As Integer '返回辅助线的序号
Dim i%
Dim in_coord As POINTAPI
Dim out_coord As POINTAPI
Dim temp_aid_line As display_line_data
in_coord.X = in_x
in_coord.Y = in_y
Dim is_ty As Byte
For i% = 1 To last_aid_line
  is_ty = is_point_on_line(in_coord, _
                        m_poi(aid_line.end_point(0)).data(0).data0.coordinate(0), _
                           m_poi(aid_line.end_point(0)).data(0).data0.coordinate(1), _
                      out_coord, _
                        temp_aid_line.end_point_coord(0), _
                         temp_aid_line.end_point_coord(1), _
                          aid_condition)
                          out_x = out_coord.X
                           out_y = out_coord.Y
  'If m_backup_line_data(0).data.visible = 1 Then
  If is_ty > 0 Then
     read_line = True
     If is_ty = 3 Then '在直线外
                out_x = out_coord.X
                   out_y = out_coord.Y
                    read_aid_line0 = True
          If (aid_line.end_point(0) > 0 And aid_line.end_point(1) > 0) Then
                  If aid_line.display = 2 Then
                   Call display_aid_line(m_object, aid_line.display_aid_line)
                  End If
                  If aid_line.display = 0 Then
                   aid_line.display_aid_line = temp_aid_line
                   aid_line.display_aid_line.total_color = fill_color
                   aid_line.display_aid_line.display = 1
                   Call display_aid_line(m_object, aid_line.display_aid_line)
                  End If
           End If
      Else
           If (aid_line.end_point(0) > 0 And aid_line.end_point(1) > 0) Then
                If aid_line.display_aid_line.display = 2 Then
                   Call display_aid_line(m_object, iad_line.display_aid_line)
                End If
           End If
      End If
   Else
           If aid_line.end_point(0) > 0 And aid_line.end_point(1) > 0 Then
                If aid_line.display = 2 Then
                   Call display_aid_line(m_object, aid_line.display_aid_line)
                End If
           End If
            read_aid_line0 = False
   End If


End Function

Public Sub init(ob As Object)
is_set_print_data = False
Call set_object(ob)
last_draw_point = 0
End Sub
Public Property Get draw_statue()
draw_statue = m_draw_statue
End Property
Public Sub set_draw_statue(st As Byte)
  m_draw_statue = st
End Sub
Public Sub set_m_line_data0(ByVal line_no%, X As Long, Y As Long, Optional draw_ty As Byte = 0)
Dim tn%
tn% = from_line_no_to_index(line_no%)
    Call m_line.item(tn%).set_m_line_data0(line_no%, X, Y, draw_ty)
End Sub
Public Sub m_draw_point()
Dim i%
For i% = 1 To m_point.Count
 m_point.item(i%).draw_m_poi_ ' (Draw_form)
Next i%
End Sub
Public Function from_point_no_to_index(ByVal point_no%) As Integer
Dim i%
from_point_no_to_index = from_point_no_to_index0(point_no%)
If from_point_no_to_index = 0 Then
 Dim m_dis_point As New display_point
  m_point.Add m_dis_point
    Call m_point.item(m_point.Count).set_object(m_object, me_class)
    from_point_no_to_index = m_point.Count
     m_point.item(m_point.Count).m_point_no = point_no%
End If
End Function
Private Function from_point_no_to_index0(ByVal point_no%) As Integer
Dim i%
For i% = 1 To m_point.Count
 If m_point.item(i%).m_point_no = point_no% Then
    from_point_no_to_index0 = i%
    Exit Function
 End If
Next i%
End Function

Private Function from_line_no_to_index0(ByVal line_no%, Optional ty As Integer = 0) As Integer 'ty=0 line ty=1 aid_line
Dim i%
If line_no% = 0 Then
    from_line_no_to_index0 = 0
Else
For i% = 1 To m_line.Count
 If (m_line.item(i%).get_line_no = line_no% And ty = 0) Then
   from_line_no_to_index0 = i%
    Exit Function
 End If
Next i%
End If
End Function
Private Function from_line_no_to_index(ByVal line_no%, Optional ty As Integer = 0) As Integer
from_line_no_to_index = from_line_no_to_index0(line_no%, ty)
If from_line_no_to_index = 0 And line_no% > 0 Then
     Dim m_dis_line As New display_line
    m_line.Add item:=m_dis_line, Key:=CStr(line_no%)
     'm_line.item(m_line.Count).line_no = line_no%
    Call m_line.item(m_line.Count).set_object(m_object, me_class)
     from_line_no_to_index = m_line.Count
    Call m_line.item(from_line_no_to_index).set_m_line_no(line_no%)
ElseIf line_no% = 0 Then
    from_line_no_to_index = 0
End If
End Function

Private Function from_circle_no_to_index(ByVal circle_no%) As Integer
from_circle_no_to_index = from_circle_no_to_index0(circle_no%)
If from_circle_no_to_index = 0 Then
  Dim m_dis_circle As New display_circle
  m_circle.Add m_dis_circle
   Call m_circle.item(m_circle.Count).set_object(m_object, me_class)
   from_circle_no_to_index = m_circle.Count
End If
End Function
Private Function from_circle_no_to_index0(ByVal circle_no%) As Integer
Dim i%
For i% = 1 To m_circle.Count
 If m_circle.item(i%).circle_no = circle_no% Then
   from_circle_no_to_index0 = i%
    Exit Function
 End If
Next i%
End Function
Public Sub set_m_point_color(ByVal point_no%, color As Byte)
Dim tn%
tn% = from_point_no_to_index0(point_no%)
If tn% > 0 Then
Call m_point.item(tn%).set_m_color(color)
End If
End Sub
Public Sub set_m_point_name(ByVal point_no%, na$)
Dim tn%
'Call C_icon.init
tn% = from_point_no_to_index(point_no%)
Call m_point.item(tn%).Set_m_name(na$)
End Sub
Public Sub set_m_point_visible(ByVal point_no%, vi As Byte)
Dim tn%
tn% = from_point_no_to_index(point_no%)
Call m_point.item(tn%).set_m_visible(vi)
End Sub
Public Sub set_m_line_visible(ByVal line_no%, vi As Byte)
Dim tn%
If run_type = 0 Or run_type > 5 Then
tn% = from_line_no_to_index0(line_no%)
If tn% = 0 And vi = 1 Then
 Dim m_dis_line As New display_line
  m_line.Add m_dis_line
    Call m_line.item(m_line.Count).set_object(m_object, me_class)
     tn% = m_line.Count
    Call m_line.item(tn%).set_m_type(condition)
    Call m_line.item(tn%).set_m_line_no(line_no%)
    Call m_line.item(tn%).set_m_line_data0(line_no%)
ElseIf tn% > 0 And vi = 0 Then
    Call m_line.Remove(tn%)
End If
End If
End Sub
Public Sub set_m_point_coordinate(ByVal point_no%, ByVal X&, ByVal Y&)
Dim tn%
If m_poi(point_no%).data(0).is_change Then
tn% = from_point_no_to_index0(point_no%)
If tn% > 0 Then
Call m_point.item(tn%).Set_m_coordinate(X&, Y&)
End If
End If
End Sub
Public Sub remove_point(point_no%)
Dim ind%
ind% = from_point_no_to_index(point_no%)
If ind% > 0 Then
 Call C_display_wenti.delete_wenti(0, point_no%, 0, 0, 0)
 Call m_point(ind%).Set_visible(0)
 Call m_point.Remove(ind%)
End If
End Sub
Public Sub remove_circle(circle_no%)
Dim ind%
ind% = from_circle_no_to_index0(circle_no%)
If ind% > 0 Then
 Call m_circle.Remove(ind%)
End If
End Sub

Public Sub set_m_circle_radii(ByVal circle_no%, ByVal radii)
If radii > 0 Then
 Call m_circle.item(from_circle_no_to_index(circle_no%)).set_m_radii(radii)
End If
End Sub
Public Sub set_m_circle_data0(ByVal circle_no%)
Dim tn%
 tn% = from_circle_no_to_index(circle_no%)
  Call m_circle.item(tn).set_m_data0(circle_no%)
 End Sub
Public Sub set_m_circle_visible(ByVal circle_no%, vi As Byte)
Call m_circle.item(from_circle_no_to_index(circle_no%)).set_m_visible(vi)
End Sub
Public Sub draw_red_point(ByVal point_no%)
 m_poi(point_no%).data(0).data0.color = conclusion_color
 Call m_point.item(from_point_no_to_index(point_no%)).set_m_color(conclusion_color)
End Sub
Public Sub redraw_point(point_no%)
Dim i%
If point_no% = 0 Then
i% = m_point.Count
Do While i% > 0
If m_point.item(i%).m_point_no < 90 Then
 Call m_point.item(i%).set_m_color(9)
Else
   m_point.Remove (i%)
End If
i% = i% - 1
Loop
Else
i% = from_point_no_to_index0(point_no%)
If i% > 0 Then
 Call m_point.item(i%).set_m_color(9)
End If
End If
End Sub
Public Sub redraw_circle(circle_no%, Optional is_recover As Boolean = False)
Dim i%
If circle_no% = 0 Then
i% = m_circle.Count
Do While i% > 0
 If m_circle.item(i%).ty = aid_condition Then
   Call m_circle.Remove(i%)
 ElseIf m_circle.item(i%).ty = conclusion Then
  Call m_circle.item(i%).set_m_color(conclusion_color)
 ElseIf m_circle.item(i%).ty = condition Then
  Call m_circle.item(i%).set_m_color(condition_color)
 End If
 i% = i% - 1
Loop
Else
 i% = from_circle_no_to_index0(circle_no%)
 If i% > 0 Then
 If is_recover Then
   Call m_circle.item(i%).redraw(is_recover)
 Else
 If m_circle.item(i%).ty = aid_condition Then
   Call m_circle.Remove(i%)
 ElseIf m_circle.item(i%).ty = conclusion Then
  Call m_circle.item(i%).set_m_color(conclusion_color)
 ElseIf m_circle.item(i%).ty = condition Then
  Call m_circle.item(i%).set_m_color(condition_color)
 End If
 End If
End If
End If
End Sub

Public Sub set_dot_line(ByVal p1%, ByVal p2%, ByVal l%, aid_point_no%)
Dim i%, no%
Dim tn(1) As Integer
Dim t_line As line_data0_type
If p1% = 0 And p2% = 0 Then
 p1% = m_lin(l%).data(0).data0.poi(0)
 p2% = m_lin(l%).data(0).data0.poi(1)
Else
l% = line_number0(p1%, p2%, tn(0), tn(1))
If tn(0) > tn(1) Then
Call exchange_two_integer(p1%, p2%)
End If
End If
 no% = from_line_no_to_index(l%)
 If aid_point_no% > 0 Then
 Call m_line.item(no%).set_m_dot_line(aid_point_no%)
 ElseIf m_lin(l%).data(0).data0.in_point(0) = 2 And m_lin(l%).data(0).data0.visible = 0 Then
 Call m_line.item(no%).set_m_dot_line_(p1%, p2%)
 End If
End Sub

Public Sub set_object(ob As Object)
Dim i%
Set m_object = ob
For i% = 1 To m_point.Count
Call m_point.item(i%).set_object(ob, me_class)
Next i%
For i% = 1 To m_aid_line.Count
Call m_aid_line.item(i%).set_object(ob, me_class)
Next i%
For i% = 1 To m_line.Count
Call m_line.item(i%).set_object(ob, me_class)
Next i%
For i% = 1 To m_circle.Count
Call m_circle.item(i%).set_object(ob, me_class)
Next i%
End Sub
Public Sub flash_point(ByVal p%)
Dim tn%
On Error GoTo flash_point_out
yidian_stop = False
tn% = from_point_no_to_index(p%)
Draw_form.Timer1.interval = 300
Draw_form.Timer1.Enabled = True '启动flash_point
'Call C_icon.set_m_is_display(True)
MDIForm1.Timer1.Enabled = True '启动光标
Do
  Do
  DoEvents
  Loop Until draw_time_act = True
   draw_time_act = False
    If red_or_black = False Then
    Call m_point.item(tn%).set_m_color_(conclusion_color)
     red_or_black = True
   Else
    Call m_point.item(tn%).set_m_color_(condition_color)
     red_or_black = False
   End If
Loop Until yidian_stop = True
flash_point_out:
Draw_form.Timer1.Enabled = False
'Call C_icon.set_m_is_display(False)
MDIForm1.Timer1.Enabled = False '关闭光标
End Sub
Public Sub open_flash_point(point_no%)

End Sub
Public Sub set_m_circle_color(circle_no%, color As Byte)
Dim tn%
tn% = from_circle_no_to_index(circle_no%)
Call m_circle.item(tn%).set_m_color(color)
End Sub

Public Sub set_m_circle_type(circle_no%, ty As Byte)
Dim tn%
tn% = from_circle_no_to_index(circle_no%)
Call m_circle.item(tn%).set_m_type(ty, tn%)
circle_no% = tn%
End Sub

Public Sub delete_circle(circle_no%)
Dim tn%
tn% = from_circle_no_to_index(circle_no%)
Call m_circle.Remove(tn%)
End Sub

Public Sub set_circle_center(ByVal circle_no%, ByVal c_center%)
Dim tn%
tn% = from_circle_no_to_index(circle_no%)
Call m_circle.item(tn%).set_center(c_center)
End Sub

Public Sub delete_line(line_no%)
Dim tn%
tn% = from_line_no_to_index(line_no%)
Call m_line.Remove(tn%)
End Sub
Public Sub delete_point(ByVal point_no%)
Dim tn%
If tn% > 0 Then
tn% = from_point_no_to_index0(point_no%)
Call m_point.Remove(tn%)
End If
End Sub
Public Sub picture_scale()
Dim i%
If is_set_print_data = False Then
is_set_print_data = True
m_picture_top_left_x = m_poi(1).data(0).data0.coordinate.X
m_picture_top_left_y = m_poi(1).data(0).data0.coordinate.Y
m_picture_bottom_right_x = m_poi(1).data(0).data0.coordinate.X
m_picture_bottom_right_y = m_poi(1).data(0).data0.coordinate.Y
For i% = 2 To last_conditions.last_cond(1).point_no
m_picture_top_left_x = min_for_long(m_picture_top_left_x, m_poi(i%).data(0).data0.coordinate.X)
m_picture_top_left_y = min_for_long(m_picture_top_left_y, m_poi(i%).data(0).data0.coordinate.Y)
m_picture_bottom_right_x = max_for_long(m_picture_bottom_right_x, m_poi(i%).data(0).data0.coordinate.X)
m_picture_bottom_right_y = max_for_long(m_picture_bottom_right_y, m_poi(i%).data(0).data0.coordinate.Y)
Next i%
For i% = 1 To m_circle.Count
m_picture_top_left_x = min_for_long(m_picture_top_left_x, _
                              m_circle.item(i%).top_left_x)
m_picture_top_left_y = min_for_long(m_picture_top_left_y, _
                              m_circle.item(i%).top_left_y)
m_picture_bottom_right_x = max_for_long(m_picture_bottom_right_x, _
                              m_circle.item(i%).bottom_right_x)
m_picture_bottom_right_y = max_for_long(m_picture_bottom_right_y, _
                              m_circle.item(i%).bottom_right_y)
Next i%
picture_width_height_x = m_picture_bottom_right_x - m_picture_top_left_x
picture_width_height_y = m_picture_bottom_right_y - m_picture_top_left_y
picture_range = max_for_long(picture_width_height_x, picture_width_height_y)
End If
End Sub

Public Sub print_picture(ob As Object, ByVal start_x&, ByVal start_y&, ty As Byte, _
          ty1 As Byte)
'用于重画图 ty -ob ,ty1 position
Dim i%
Call picture_scale '确定原图大小
'Call trans_data_to_print
'ob.FontSize = 10
ob.FontSize = 14
If ty = 0 Then
ob.DrawWidth = 1
Else
ob.DrawWidth = 2
End If
ob.DrawStyle = 0
'picture_top_left_x = start_x&
'picture_top_left_y = start_y&
For i% = 1 To m_point.Count
  Call m_point.item(i%).picture_orig_x(start_x&)
  Call m_point.item(i%).picture_orig_y(start_y&)
  Call m_point.item(i%).picture_top_left_x(m_picture_top_left_x)
  Call m_point.item(i%).picture_top_left_y(m_picture_top_left_y)
  Call m_point.item(i%).picture_range(picture_range)
Next i%
For i% = 1 To m_circle.Count
  Call m_circle.item(i%).picture_orig_x(start_x&)
  Call m_circle.item(i%).picture_orig_y(start_y&)
  Call m_circle.item(i%).picture_top_left_x(m_picture_top_left_x)
  Call m_circle.item(i%).picture_top_left_y(m_picture_top_left_y)
  Call m_circle.item(i%).picture_range(picture_range)
Next i%
For i% = 1 To m_line.Count
  Call m_line.item(i%).picture_orig_x(start_x&)
  Call m_line.item(i%).picture_orig_y(start_y&)
  Call m_line.item(i%).picture_top_left_x(m_picture_top_left_x)
  Call m_line.item(i%).picture_top_left_y(m_picture_top_left_y)
  Call m_line.item(i%).picture_range(picture_range)
Next i%
For i% = 1 To m_point.Count
 Call m_point.item(i%).print_(ob)
Next i%
For i% = 1 To m_line.Count
 Call m_line.item(i%).print_(ob)
Next i%
For i% = 1 To m_circle.Count
 Call m_circle.item(i%).print_(ob)
Next i%

End Sub

Public Function picture_top_left_x() As Long
picture_top_left_x = m_picture_top_left_x
End Function
Public Function picture_top_left_y() As Long
picture_top_left_y = m_picture_top_left_y
End Function
Public Function picture_bottom_right_x() As Long
picture_bottom_right_x = m_picture_bottom_right_x
End Function
Public Function picture_bottom_right_y() As Long
picture_bottom_right_y = m_picture_bottom_right_y
End Function
Public Function chose_point_name_position(left_top_x As Long, left_top_y As Long) As Boolean
Dim l_t As POINTAPI
Dim i%
l_t.X = left_top_x  'print_P(p_no%).coordinate.X + left_top.X
l_t.Y = left_top_y  'print_P(p_no%).coordinate.Y + left_top.Y
For i% = 1 To m_line.Count
 If m_line.item(i%).is_char_non_interset_line(l_t.X, l_t.Y) = False Then
           chose_point_name_position = False
            Exit Function
 End If
Next i%
For i% = 1 To m_circle.Count
 If m_circle.item(i%).is_char_non_interset_circle(l_t.X, l_t.Y) = False Then
           chose_point_name_position = False
            Exit Function
 End If
Next i%
           chose_point_name_position = True
End Function

Public Sub set_me_class(ob As Object)
Set me_class = ob
End Sub


Public Function m_last_draw_point() As Integer
m_last_draw_point = last_draw_point
End Function
Public Sub get_circ_data(draw_circle_no%)
Dim ind%
 ind% = from_circle_no_to_index(draw_circle_no%)
  Call m_circle.item(ind%).Get_data
End Sub
Public Function read_m_circle(ByVal X As Long, Y As Long, point_no%, ty As Byte) As Integer
Dim i%
Dim r!
Dim c_n%
For i% = 1 To m_circle.Count
If is_point_in_points(point_no%, m_Circ(i%).data(0).data0.in_point) > 0 Then
    GoTo mark1
End If
   c_n% = m_circle.item(i%).read_circle(X, Y)
   If c_n% <> 0 Then
      read_m_circle = c_n%
       Exit Function
   End If
mark1:
Next i%
End Function
Public Sub remove_aid_circle()
Dim i%, j%
For i% = m_circle.Count To 1 Step -1
     If m_circle.item(i%).type = aid_condition Then
      m_circle.Remove (i%)
     End If
Next i%
End Sub
Public Function circle_no(ByVal circ_no%, is_draw As Boolean) As Integer
If circ_no% >= 0 Then
   circle_no = circ_no%
Else
   circ_no% = from_circle_no_to_index(circ_no%)
 '  Call m_circle.item(circ_no%).set_
End If
End Function
Public Function is_old_m_circle() As Integer
Dim i%, n%
For i% = 1 To m_circle.Count
 n% = m_circle.item(i%).is_old_circle
  If n% <> 0 Then
     is_old_m_circle = n%
      Exit Function
  End If
Next i%
End Function
Public Sub redraw()
Dim i%, line_no%
m_last_aid_line = 0
For i% = m_line.Count To 1 Step -1
 If m_line.item(i%).kill_aid_line Then
 Call m_line.item(i%).redraw(1) '画线结束
   m_lin(i%).data(0).is_change = 0
     'Call m_line.Remove(i%)
 End If
Next i%
For i% = 1 To m_circle.Count
 Call m_circle.item(i%).redraw
Next i%
End Sub
Public Sub PointCoordinateChange(point_no As Integer)
Dim i%
Dim tn%
tn% = from_point_no_to_index(point_no)
m_point.item(tn%).draw_m_poi_ '重画点
'**************************************
If m_poi(point_no).data(0).is_change Then
  For i% = 1 To m_circle.Count
   Call m_circle.item(i%).PointCoordinateChange(point_no%)
  Next i%
For i% = 1 To m_line.Count
   Call m_line.item(i%).PointCoordinateChange(point_no%)
Next i%
End If
End Sub
Public Sub set_circle_radii_no(ByVal circle_no%, ByVal radii_no%)
Dim tn%
tn% = from_circle_no_to_index0(circle_no%)
m_circle.item(tn%).set_radii_no (radii_no%)
End Sub
Public Sub init_change()
Dim i%
For i% = 1 To last_conditions.last_cond(1).point_no
m_poi(i%).data(0).is_change = False
Next i%
For i% = 1 To last_conditions.last_cond(1).line_no
m_lin(i%).data(0).is_change = 0
Next i%
For i% = 1 To C_display_picture.m_circle.Count
Call C_display_picture.m_circle.item(i%).is_change(False)
Next i%
'For i% = 1 To last_conditions.last_cond(1).aid_circle_no
'm_Circ(-i%).data(0).is_change = False
'Next i%
End Sub
Public Sub draw_circle(circle_no%, X As Long, Y As Long, Optional color_no As Integer = -1)
Dim tn%
tn% = from_circle_no_to_index0(circle_no%)
If tn% > 0 Then
Call m_circle.item(tn%).draw_circle_(X, Y, color_no%)
End If
End Sub
Public Sub CircleChange(ByVal circle_no%)
Dim i%
For i% = 1 To m_point.Count
 Call m_point.item(i%).CircleChange(circle_no%)
Next i%
End Sub

Public Sub draw_point(ByVal point_no%)
Dim tn%

tn% = from_point_no_to_index(point_no%)
m_point.item(tn%).draw_m_poi_
m_poi(point_no%).data(0).is_change = False
End Sub

Public Sub draw_line(line_no%, X As Long, Y As Long, Optional draw_type As Byte = 0)
Dim tn%
 tn% = from_line_no_to_index0(line_no%)
If tn% > 0 Then
 Call m_line.item(tn%).draw_line_(X, Y, draw_type)
End If
End Sub
Public Sub Get_line_point_coordinate(line_no%, point_no%, X&, Y&)
Dim tn%
 tn% = from_line_no_to_index0(line_no%)
If tn% > 0 Then
 Call m_line.item(tn%).Get_line_point_coordinate(point_no%, X&, Y&)
End If
End Sub
Public Sub draw_aid_line(ByVal l1%, ByVal l2%, ByVal X As Long, ByVal Y As Long)
Dim tn(1) As Integer
Dim tn_(1) As Integer
tn(0) = from_line_no_to_index(l1%)
tn(1) = from_line_no_to_index(l2%)
If tn(0) > 0 And tn(1) > 0 And aid_line(0) = tn(0) And aid_line(1) = tn(1) Then
  Exit Sub
Else
  tn_(0) = from_line_no_to_index(aid_line(0))
  tn_(1) = from_line_no_to_index(aid_line(1))
  If tn_(0) > 0 Then
    Call m_line.item(tn(0)).read_line(X, Y, 0, 0, 0, 0)
  End If
  If tn_(1) > 0 Then
    Call m_line.item(tn(1)).read_line(X, Y, 0, 0, 0, 0)
  End If
  '******************************************
  'aid_line(0) = l1%
  'aid_line(1) = l2%
  'If tn(0) > 0 Then
  'Call m_line.item(tn(0)).draw_aid_line(X, Y)
  'End If
  'If tn(1) > 0 Then
  'Call m_line.item(tn(1)).draw_aid_line(X, Y)
  Call UpdateWindow(m_object.hwnd)
  End If
End Sub
Sub set_aid_line_start_point(ByVal p%, ByVal line_no%)
   Call m_line.item(line_no%).set_aid_line_start_point(p%)
End Sub
Function read_line(ByVal last_line%, in_x As Long, in_y As Long, point_no%, _
           out_x As Long, out_y As Long, is_set_data As Boolean, Optional ty As Byte = 0) As Integer
Dim i%
Dim out2_x As Long, out1_x As Long
Dim out2_y As Long, out1_y As Long
'Dim paral_or_verti As Integer
Dim line_no%
Dim temp_p(1) As Integer
Dim temp_coord(1) As POINTAPI
Dim ele1 As condition_type
Dim ele2 As condition_type
Dim t_paral_or_verti As Integer
For i% = last_line% To 1 Step -1 'm_line.Count
If is_set_data Then
   If is_point_in_points(point_no%, m_lin(i%).data(0).data0.in_point) > 0 Then '输入点已在线上
      GoTo mark1
   End If
End If
If m_line.item(i%).read_line(in_x, in_y, point_no%, out_x, out_y, _
                    out1_x, out1_y, out2_x, out2_y, t_paral_or_verti, ty) Then
    line_no% = i%
    If get_control_data(line_, line_no%, ty) = False Then
       GoTo mark1
    End If
    If is_set_data Then
    If t_paral_or_verti = 0 Then
    read_line = i%
        m_lin(i%).data(0).data0.end_point_coord(0) = _
                                           m_poi(m_lin(i%).data(0).data0.in_point(1)).data(0).data0.coordinate
    If m_lin(i%).data(0).data0.in_point(0) > 1 Then
        m_lin(i%).data(0).data0.end_point_coord(1) = _
          m_poi(m_lin(i%).data(0).data0.in_point(m_lin(i%).data(0).data0.in_point(0))).data(0).data0.coordinate
    Else
        m_lin(i%).data(0).data0.poi(0) = m_lin(i%).data(0).data0.in_point(1)
        m_lin(i%).data(0).data0.end_point_coord(1).X = out2_x
        m_lin(i%).data(0).data0.end_point_coord(1).Y = out2_y
        m_lin(i%).data(0).data0.total_color = condition_color
        m_lin(i%).data(0).data0.type = condition
        Call m_line.item(i%).redraw
    End If
    ElseIf t_paral_or_verti = verti_ Or t_paral_or_verti = paral_ Then
        paral_or_verti = t_paral_or_verti
        temp_coord(0).X = out1_x
        temp_coord(0).Y = out1_y
        If t_paral_or_verti = paral_ Then
           temp_coord(1) = add_POINTAPI(temp_coord(0), minus_POINTAPI(second_end_point_coordinate(i%), _
                 m_poi(m_lin(i%).data(0).data0.depend_poi(0)).data(0).data0.coordinate))
        Else
           temp_coord(1) = add_POINTAPI(temp_coord(0), verti_POINTAPI(minus_POINTAPI(second_end_point_coordinate(i%), _
                 m_poi(m_lin(i%).data(0).data0.depend_poi(0)).data(0).data0.coordinate)))
        End If
        'temp_coord(1).X = out2_x
        'temp_coord(1).Y = out2_y
        temp_p(0) = m_point_number(temp_coord(0), condition, 1, condition_color, "", ele1, ele2, 0, True)
        read_line = line_number(temp_p(0), 0, temp_coord(0), temp_coord(1), _
                        depend_condition(point_, temp_p(0)), _
                         depend_condition(line_, i%), condition, condition_color, 0, 0)
        m_lin(read_line).data(0).parent.inter_type = paral_or_verti
        temp_line(draw_line_no) = read_line
                   draw_line_no = draw_line_no + 1
                       m_lin(read_line).data(0).data0.end_point_coord(0) = temp_coord(0)
                       m_lin(read_line).data(0).data0.end_point_coord(1) = temp_coord(1)
                       Call m_line.item(i%).set_aid_line_start_point(0)
    End If
    End If
    Exit Function
  End If
mark1:
Next i%
End Function
Public Function set_aid_line(head_point%, x1 As Long, y1 As Long) As Boolean
Dim i%
For i% = 1 To m_aid_line.Count
   If m_aid_line.item(i%).is_me(head_point%, x1, y1) Then
      GoTo set_aid_line_mark1
   End If
Next i%
set_aid_line_mark1:
Dim aid_line_data_ As aid_line
m_aid_line.Add (aid_line_data0)
Call m_aid_line.item(m_aid_line.Count).set_aid_line(head_point%, x1, y1)
End Function
Public Function set_aid_line_to_exist_line(x1 As Long, y1 As Long, line_no%) As Boolean
Dim i%
 i% = from_line_no_to_index(line_no%)
 If i% = 0 Then
    set_aid_linet_to_exist_line = False
 Else
   Call m_line.item(i%).set_aid_line(x1, y1)
 End If
End Function
Public Sub get_end_point(line_no%, p1%, p2%, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
 Call m_line.item(from_line_no_to_index(line_no%)).get_end_point( _
                p1%, p2%, x1, y1, x2, y2)
End Sub
Public Sub get_line_depend_data(line_no%, d_ty1 As Integer, d_no1%, _
                                         d_ty2 As Byte, d_no2%)
 Call m_line.item(from_line_no_to_index(line_no%)).get_depend_data( _
                                             d_ty1, d_no1%, d_ty2, d_no2%)
End Sub
Public Sub re_draw_line(line_no%, Optional draw_type As Integer = 0) '=0画线 =1 画线结束，恢复=2
If m_lin(line_no%).data(0).is_change = 255 Or _
    change_picture_start_no.is_picture_change Then
 Call m_line.item(from_line_no_to_index(line_no%)).redraw(draw_type)
 m_lin(line_no%).data(0).is_change = 0
End If
End Sub

Private Sub Class_Terminate()
Set m_line = Nothing
Set m_point = Nothing
Set m_circle = Nothing
Set m_conclusion_line = Nothing
Set m_conclusion_circle = Nothing
Set m_aid_line = Nothing
End Sub

Public Sub set_m_line_color(line_no%, color_no As Byte)
 Call m_line.item(from_line_no_to_index(line_no%)).set_m_color(color_no)
End Sub

Public Sub redraw_line(ByVal line_no%)
Dim tn%
 tn% = from_line_no_to_index0(line_no%)
If tn% > 0 Then
 Call m_line.item(tn%).redraw_line
End If
 
End Sub

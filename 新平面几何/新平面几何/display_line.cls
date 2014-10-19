VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "display_line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*********************************
'备份直线数据,仅作重画图用
Private Type backup_line_data_type
data As line_data0_type
end_poi(1) As Integer '可视端点的序号
display_line0 As display_line_data
aid_line As aid_line_data
total_section_no As Integer  '分段
display_line(1 To 8) As display_line_data '每一段的数据
is_dot As Boolean
End Type
'Private data As line_data0_type
'**************************************
Private is_draw_aid_line As Boolean '是否是辅助线
Private m_draw_type As Byte '虚线,实线
Private m_line_no As Integer '直线序号
Private m_depend_condition(1) As condition_type '
Private m_backup_line_data(1)  As backup_line_data_type
Private m_aid_point(1) As POINTAPI
Private m_object As Object
Private m_picture_range As Long
Private m_picture_orig As POINTAPI
'Private m_picture_orig_y As Long
Private m_picture_top_left As POINTAPI
'Private m_picture_top_left_y As Long
Private is_set_print_data As Boolean
Private father_class As Object
Public Sub picture_orig_x(p_or_x As Long)
  m_picture_orig.X = p_or_x
End Sub
Public Sub picture_orig_y(p_or_y As Long)
  m_picture_orig.Y = p_or_y
End Sub
Public Sub picture_range(p_r As Long)
  m_picture_range = p_r
  Call Backup(1, 0, 0)
End Sub
Public Sub picture_top_left_x(top_x As Long)
  m_picture_top_left.X = top_x
End Sub
Public Sub picture_top_left_y(top_y As Long)
  m_picture_top_left.Y = top_y
End Sub

Property Get line_no() As Integer
line_no = m_line_no
End Property

Public Sub Get_line_point_coordinate(no%, X&, Y&)
X& = m_backup_line_data(0).display_line0.end_point_coord(no%).X
Y& = m_backup_line_data(0).display_line0.end_point_coord(no%).Y
End Sub
Public Function get_line_no() As Integer
 get_line_no = m_line_no
End Function
Public Sub set_m_line_data0(line_no%, X As Long, Y As Long, Optional draw_ty As Byte = 0)
Dim i%
   Call draw_line_(X, Y, draw_ty)
   m_lin(line_no%).data(0).is_change = False
End Sub
Private Sub draw_line(ty As Byte, Optional draw_ty As Byte = 0)   '12.30 ty=1  0
Dim i%
If m_backup_line_data(draw_ty).data.visible > 0 Then
For i% = 1 To m_backup_line_data(draw_ty).total_section_no
If is_different_display_line(m_backup_line_data(0).display_line(i%), m_backup_line_data(1).display_line(i%)) Then
   Call Drawline(m_object, QBColor(m_backup_line_data(draw_ty).display_line(i%).color), 0, _
                  m_backup_line_data(draw_ty).display_line(i%).end_point_coord(0), _
                   m_backup_line_data(draw_ty).display_line(i%).end_point_coord(1), _
                    m_backup_line_data(draw_ty).data.in_point(10), m_backup_line_data(draw_ty).data.visible)
 End If
 Next i%
 End If
End Sub
Private Function is_different_display_line(dis_line1 As display_line_data, dis_line2 As display_line_data) As Boolean
 If dis_line1.color = dis_line2.color Then
  If dis_line1.end_point_coord(0).X = dis_line2.end_point_coord(0).X And _
       dis_line1.end_point_coord(0).Y = dis_line2.end_point_coord(0).Y Then
   If dis_line1.end_point_coord(1).X = dis_line2.end_point_coord(1).X And _
        dis_line1.end_point_coord(1).Y = dis_line2.end_point_coord(1).Y Then
     is_different_display_line = False
      Exit Function
   End If
  End If
 End If
     is_different_display_line = True
End Function
Public Sub set_m_visible(vi As Byte)
If m_backup_line_data(1).data.visible <> vi Then
   m_lin(m_line_no%).data(0).data0.visible = vi
    Call draw_line_(0, 0)
End If
End Sub
Public Sub set_m_color(co As Byte)
If m_backup_line_data(1).data.total_color <> co Then
   m_lin(m_line_no%).data(0).data0.total_color = co
    m_lin(m_line_no%).data(0).is_change = True
     Call draw_line_(0, 0)
End If
End Sub
Public Sub set_m_type(ty As Byte)
If m_backup_line_data(1).data.type <> ty Then
   m_lin(m_line_no%).data(0).data0.type = ty
    Call draw_line_(0, 0)
End If
End Sub
Public Sub put_arrow_to_line_point(color As Long, draw_ty As Byte)
Dim i%, tn1%, tn2% 'ty=1 消去原来图象
Dim dire As Integer
If regist_data.run_type = 1 Then
If compare_two_point(m_poi(m_backup_line_data(draw_ty).data.poi(0)).data(0).data0.coordinate, _
       m_poi(m_backup_line_data(draw_ty).data.poi(1)).data(0).data0.coordinate, 0, 0, 6) = 1 Then
   dire = m_backup_line_data(draw_ty).data.in_point(10)
Else
   dire = -m_backup_line_data(draw_ty).data.in_point(10)
End If
If dire = 0 Then
 Exit Sub
ElseIf dire = 1 Then
 For i% = m_backup_line_data(draw_ty).end_poi(0) + 1 To m_backup_line_data(draw_ty).end_poi(1) - 1
        Call draw_arrow(m_object, m_poi(Abs(m_backup_line_data(draw_ty).data.in_point(m_backup_line_data(draw_ty).end_poi(0)))).data(0).data0.coordinate.X, _
                    m_poi(Abs(m_backup_line_data(draw_ty).data.in_point(m_backup_line_data(draw_ty).end_poi(0)))).data(0).data0.coordinate.Y, _
                     m_poi(Abs(m_backup_line_data(draw_ty).data.in_point(i%))).data(0).data0.coordinate.X, _
                       m_poi(Abs(m_backup_line_data(draw_ty).data.in_point(i%))).data(0).data0.coordinate.Y, _
                         color)
 Next i%
ElseIf dire = -1 Then
 For i% = m_backup_line_data(draw_ty).end_poi(0) + 1 To m_backup_line_data(draw_ty).end_poi(1) - 1
    Call draw_arrow(m_object, m_poi(Abs(m_backup_line_data(draw_ty).data.in_point(m_backup_line_data(draw_ty).end_poi(1)))).data(0).data0.coordinate.X, _
                         m_poi(Abs(m_backup_line_data(draw_ty).data.in_point(m_backup_line_data(draw_ty).data.poi(1)))).data(0).data0.coordinate.Y, _
                          m_poi(Abs(m_backup_line_data(draw_ty).data.in_point(i%))).data(0).data0.coordinate.X, _
                           m_poi(Abs(m_backup_line_data(draw_ty).data.in_point(i%))).data(0).data0.coordinate.Y, color)
 Next i%
End If
End If
End Sub
Public Sub set_m_line_no(line_no%)
m_line_no = line_no%
End Sub
Public Sub set_m_aid_line_no(aid_line_no%)
m_aid_line_no = aid_line_no%
End Sub

Public Function Backup(ty As Integer, X As Long, Y As Long) As Boolean '=0,画先,=1恢复
Dim i%, j%
Dim is_minus As Boolean
If (ty = 0 Or ty = 1) And m_lin(m_line_no%).data(0).is_change Then
     Backup = True
 m_backup_line_data(1) = m_backup_line_data(0) '备份直线数据
 If m_backup_line_data(0).aid_line.display_aid_line.display = 2 Then
    Call display_aid_line(m_object, m_backup_line_data(0).aid_line.display_aid_line) '消除aid_line
 End If
'ElseIf ty = 0 Then '画新线
 m_backup_line_data(0).data = m_lin(m_line_no%).data(0).data0 '备份直线数据
'If m_backup_line_data(0).data.in_point(0) > 0 Then '有点
m_backup_line_data(0).end_poi(0) = 0
m_backup_line_data(0).end_poi(1) = 0
    If m_lin(m_line_no%).data(0).data0.in_point(0) = 1 Then '单点
       m_backup_line_data(0).end_poi(0) = m_lin(m_line_no%).data(0).data0.in_point(1)
         m_backup_line_data(0).display_line(1).end_point_coord(0) = m_poi(Abs(m_lin(m_line_no%).data(0).data0.in_point(1))).data(0).data0.coordinate
          m_backup_line_data(0).display_line(1).end_point_coord(1).X = X
           m_backup_line_data(0).display_line(1).end_point_coord(1).Y = Y
            m_backup_line_data(0).display_line(1).color = m_lin(m_line_no%).data(0).data0.total_color
             m_backup_line_data(0).display_line0 = m_backup_line_data(0).display_line(1)
              m_backup_line_data(0).total_section_no = 1
               m_backup_line_data(0).aid_line.end_point(0) = m_backup_line_data(0).display_line0.end_point_coord(0)
                m_backup_line_data(0).aid_line.end_point(1) = m_backup_line_data(0).display_line0.end_point_coord(1)
    ElseIf m_lin(m_line_no%).data(0).data0.visible = 1 And m_lin(m_line_no%).data(0).data0.in_point(0) > 1 And _
                m_lin(line_no%).data(0).is_change Then '数据有变化，才需重画直线
                 Backup = True
                 m_backup_line_data(0).total_section_no = 1
                  If m_lin(line_no%).data(0).data0.total_color > 0 Then
                   m_backup_line_data(0).display_line(1).color = m_lin(m_line_no%).data(0).data0.total_color
                  Else
                   m_backup_line_data(0).display_line(1).color = m_lin(m_line_no%).data(0).data0.color(1)
                  End If
                   m_backup_line_data(0).display_line(1).end_point_coord(0) = _
                      m_poi(m_lin(line_no%).data(0).data0.in_point(1)).data(0).data0.coordinate
      For i% = 2 To m_lin(m_line_no%).data(0).data0.in_point(0) - 1 'To m_backup_line_data(draw_ty).end_poi(1) - 1
       If m_lin(m_line_no%).data(0).data0.color(i%) <> m_lin(m_line_no%).data(0).data0.color(i% - 1) Then '变换颜色
           m_backup_line_data(0).display_line(m_backup_line_data(0).total_section_no).end_point_coord(1) = _
                                   m_poi(m_lin(line_no%).data(0).data0.in_point(i%)).data(0).data0.coordinate
              m_backup_line_data(0).total_section_no = m_backup_line_data(0).total_section_no + 1
                m_backup_line_data(0).display_line(m_backup_line_data(0).total_section_no).end_point_coord(0) = _
                                   m_poi(m_lin(line_no%).data(0).data0.in_point(i%)).data(0).data0.coordinate
       End If
         Next i%
           m_backup_line_data(0).display_line(m_backup_line_data(0).total_section_no).end_point_coord(1) = _
                                   m_poi(m_lin(line_no%).data(0).data0.in_point(i%)).data(0).data0.coordinate
         If m_backup_line_data(0).total_section_no = 1 And m_lin(line_no%).data(0).data0.color(1) = condition_color Then
                            m_lin(line_no%).data(0).is_change = False
         End If
            m_backup_line_data(0).end_poi(0) = m_lin(line_no%).data(0).data0.in_point(1)
            m_backup_line_data(0).end_poi(1) = m_lin(line_no%).data(0).data0.in_point(m_lin(line_no%).data(0).data0.in_point(0))
            m_backup_line_data(0).aid_line.end_point(0) = m_poi(m_backup_line_data(0).end_poi(0)).data(0).data0.coordinate
            m_backup_line_data(0).aid_line.end_point(1) = m_poi(m_backup_line_data(0).end_poi(1)).data(0).data0.coordinate

    End If
'ElseIf ty = 1 Then
 '       If m_lin(m_line_no%).data(0).data0.type = condition And m_lin(m_line_no%).data(0).data0.color <> condition_color Then
  '        m_backup_line_data(0).display_line0.end_point_coord(0) = m_poi(m_lin(m_line_no%).data(0).data0.poi(0)).data(0).data0.coordinate
   '         m_backup_line_data(0).display_line0.end_point_coord(1) = m_poi(m_lin(m_line_no%).data(0).data0.poi(0)).data(1).data0.coordinate
    '         m_backup_line_data(0).display_line0.color = condition_color
     '      m_backup_line_data(0).display_line(1) = m_backup_line_data(0).display_line0
      '      m_backup_line_data(0).total_section_no = 1
       '       m_lin(m_line_no%).data(0).data0.color = condition_color
 ElseIf ty = 2 Then '恢复
   m_backup_line_data(1) = m_backup_line_data(0)
    Backup = True
   m_lin(m_line_no).data(0).data0 = m_backup_line_data(0).data
 ElseIf ty = 3 Then
        is_set_print_data = True
     m_backup_line_data(0).display_line0.end_point_coord(0).X = (m_backup_line_data(0).display_line0.end_point_coord(0).X - _
           m_picture_top_left.X) * 2888 / m_picture_range + _
            m_picture_orig.X
     m_backup_line_data(0).display_line0.end_point_coord(0).Y = (m_backup_line_data(0).display_line0.end_point_coord(0).Y - _
           m_picture_top_left.Y) * 2888 / m_picture_range + _
            m_picture_orig.Y
     m_backup_line_data(0).display_line0.end_point_coord(1).X = (m_backup_line_data(0).display_line0.end_point_coord(1).X - _
           m_picture_top_left.X) * 2888 / m_picture_range + _
            m_picture_orig.X
     m_backup_line_data(0).display_line0.end_point_coord(1).Y = (m_backup_line_data(0).display_line0.end_point_coord(1).Y - _
           m_picture_top_left.Y) * 2888 / m_picture_range + _
            m_picture_orig.Y
End If
'*****************************************************************************************************************************
End Function


Private Sub Class_Initialize()
Dim i%
For i% = 1 To 8
m_backup_line_data(0).display_line(i%).end_point_coord(0).X = 10000
m_backup_line_data(0).display_line(i%).end_point_coord(0).Y = 10000
m_backup_line_data(0).display_line(i%).end_point_coord(1).X = 10000
m_backup_line_data(0).display_line(i%).end_point_coord(1).Y = 10000
'm_backup_line_data(0).display_line(i%).total_color = 3
m_backup_line_data(1).display_line(i%).end_point_coord(0).X = 10000
m_backup_line_data(1).display_line(i%).end_point_coord(0).Y = 10000
m_backup_line_data(1).display_line(i%).end_point_coord(1).X = 10000
m_backup_line_data(1).display_line(i%).end_point_coord(1).Y = 10000
'm_backup_line_data(1).display_line(i%).total_color = 3
Next i%
m_backup_line_data(0).display_line0.end_point_coord(0).X = 10000
m_backup_line_data(0).display_line0.end_point_coord(0).Y = 10000
m_backup_line_data(0).display_line0.end_point_coord(1).X = 10000
m_backup_line_data(0).display_line0.end_point_coord(1).Y = 10000
'm_backup_line_data(0).display_line0.total_color = 3
m_backup_line_data(1).display_line0.end_point_coord(0).X = 10000
m_backup_line_data(1).display_line0.end_point_coord(0).Y = 10000
m_backup_line_data(1).display_line0.end_point_coord(1).X = 10000
m_backup_line_data(1).display_line0.end_point_coord(1).Y = 10000
'm_backup_line_data(1).display_line0.total_color = 3
is_set_print_data = False
End Sub

Public Sub set_object(ob As Object, ob1 As Object)
Set m_object = ob
Set father_class = ob1
End Sub

Private Sub Class_Terminate()
Call remove_aid_line
   Call Drawline(m_object, QBColor(m_backup_line_data(0).display_line(1).color), 0, _
                  m_backup_line_data(0).display_line(1).end_point_coord(0), _
                   m_backup_line_data(0).display_line(1).end_point_coord(1), _
                    m_backup_line_data(0).data.in_point(10), m_backup_line_data(0).data.visible)
End Sub
Public Sub print_(ob As Object)
  Set m_object = ob
   Call Backup(3, 0, 0)
    Call draw_line(0, 1)
End Sub

Public Function is_char_non_interset_line(char_left_top_x As Long, _
                                            char_left_top_y As Long) As Boolean
Dim d(3) As Integer
Dim tp(3) As POINTAPI
Dim l_p(1) As POINTAPI
Dim i%
tp(2).X = char_left_top_x
tp(2).Y = char_left_top_y
tp(3).X = tp(2).X + 110
tp(3).Y = tp(2).Y + 140
tp(0).X = tp(2).X
tp(0).Y = tp(3).Y
tp(1).Y = tp(2).Y
tp(1).X = tp(3).X
If (m_backup_line_data(0).display_line0.end_point_coord(0).X < tp(0).X And _
      m_backup_line_data(0).display_line0.end_point_coord(1).X < tp(0).X) Or _
      (m_backup_line_data(0).display_line0.end_point_coord(0).X > tp(1).X And _
       m_backup_line_data(0).display_line0.end_point_coord(1).X > tp(1).X) Or _
   (m_backup_line_data(0).display_line0.end_point_coord(0).Y < tp(0).Y And _
     m_backup_line_data(0).display_line0.end_point_coord(1).Y < tp(0).Y) Or _
       (m_backup_line_data(0).display_line0.end_point_coord(0).Y > tp(1).Y And _
         m_backup_line_data(0).display_line0.end_point_coord(1).Y > tp(1).Y) Then
   is_char_non_interset_line = True
Else
For i% = 0 To 3
d(i%) = area_triangle_from_three_point(tp(i%), m_backup_line_data(0).display_line0.end_point_coord(0), _
                          m_backup_line_data(0).display_line0.end_point_coord(1))
Next i%
If d(0) = 1 Then
 If d(1) = 1 And d(2) = 1 And d(3) = 1 Then
  is_char_non_interset_line = True
 Else
  is_char_non_interset_line = False
 End If
ElseIf d(0) = 0 Then
  is_char_non_interset_line = False
Else
 If d(1) = -1 And d(2) = -1 And d(3) = -1 Then
  is_char_non_interset_line = True
 Else
  is_char_non_interset_line = False
 End If
End If
End If
End Function

Public Sub draw_line_(X As Long, Y As Long, Optional draw_ty As Byte = 0)
'ty=1 暂时 第一次,ty=0 永久 ty=2 数据保持不变
'Call remove_aid_line
m_draw_type = 1
If draw_ty = 0 Then
 Call Backup(0, X, Y)
 'Call draw_line(1, 1)
ElseIf draw_ty = 1 Then
 Call Backup(1, X, Y)
 'Call draw_line(1, 1)
ElseIf draw_ty = 2 And m_draw_type = 1 Then
 Call Backup(2, X, Y)
 'Call draw_line(1, 1)
Else
 Exit Sub
End If
 Call draw_line(0, 1)
 Call draw_line(0, 0)
 m_draw_type = 1
End Sub
Public Sub draw_aid_line(ByVal X As Long, ByVal Y As Long, Optional draw_ty As Byte = 1) '画辅助线，draw_ty=1
Dim t_coord As POINTAPI
Dim com_ty(1) As Integer
If m_backup_line_data(draw_ty).data.in_point(0) < 2 Then
    Exit Sub
End If
t_coord.X = X '
t_coord.Y = Y '鼠标位置
If m_backup_line_data(draw_ty).data.type <> aid_condition And _
     m_backup_line_data(draw_ty).data.in_point(0) < 2 Then '画线过程
      Exit Sub
ElseIf m_backup_line_data(draw_ty).data.type = aid_condition Then '备分直线是辅助性质
   If m_backup_line_data(draw_ty).data.poi(0) > 0 And _
        m_poi(m_backup_line_data(draw_ty).data.poi(0)).data(0).data0.visible = 1 And _
         m_backup_line_data(draw_ty).data.poi(1) = 0 Then '单端点可视直线
    m_aid_point(0) = m_backup_line_data(draw_ty).display_line0.end_point_coord(0) '单端点可视直线的已知端点
    m_aid_point(1) = t_coord '鼠标位置
   Else
    m_aid_point(0) = m_backup_line_data(draw_ty).display_line0.end_point_coord(0) '两端点可视直线
    m_aid_point(1) = m_backup_line_data(draw_ty).display_line0.end_point_coord(1)
   End If
Else '备分直线不是辅助性质
 com_ty(0) = compare_two_point(t_coord, m_backup_line_data(draw_ty).display_line0.end_point_coord(0), 0, 0, 5) '可视直线两端点与鼠标位置比较
 com_ty(1) = compare_two_point(m_backup_line_data(draw_ty).display_line0.end_point_coord(1), t_coord, 0, 0, 5)
If m_backup_line_data(draw_ty).data.visible = 1 Then '可视
   If com_ty(0) = 1 Then '鼠标位置在第一点左
    m_aid_point(0) = t_coord '取辅助线的第一端点左
    m_aid_point(1) = m_backup_line_data(draw_ty).display_line0.end_point_coord(0) '取辅助线的第二端点，取直线的左部分
   ElseIf com_ty(1) = 1 Then ''取辅助线的第一端点右
    m_aid_point(0) = m_backup_line_data(draw_ty).display_line0.end_point_coord(1)
    m_aid_point(1) = t_coord
   Else
    m_aid_point(1) = m_aid_point(0)
   End If
Else
   If com_ty(0) = 1 Then
    m_aid_point(0) = t_coord
    m_aid_point(1) = m_backup_line_data(draw_ty).display_line0.end_point_coord(0)
   ElseIf com_ty(1) = 1 Then
    m_aid_point(0) = m_backup_line_data(draw_ty).display_line0.end_point_coord(1)
    m_aid_point(1) = t_coord
   Else
    m_aid_point(0) = m_backup_line_data(draw_ty).display_line0.end_point_coord(0)
    m_aid_point(1) = m_backup_line_data(draw_ty).display_line0.end_point_coord(1)
   End If
End If
End If
Call draw_aid_line_0 '重画辅助线
End Sub
Public Sub remove_aid_line()
        If m_backup_line_data(0).aid_line.display_aid_line.display = 2 Then
          Call display_aid_line(m_object, m_backup_line_data(0).aid_line.display_aid_line)
        End If
          m_backup_line_data(0).aid_line.start_point = 0
End Sub
Sub set_aid_line_start_point(ByVal p%)
  m_backup_line_data(0).aid_line.start_point = p%
  If p% = 0 And m_backup_line_data(0).aid_line.display_aid_line.display = 2 Then
     Call display_aid_line(m_object, m_backup_line_data(0).aid_line.display_aid_line) '消除辅助线
  End If
End Sub
Public Function read_line(ByVal in_x As Long, ByVal in_y As Long, point_no%, out_x As Long, out_y As Long, _
                 out1_x As Long, out1_y As Long, out2_x As Long, out2_y As Long, _
                                                     paral_or_verti_ As Integer, Optional ty As Byte = 0) As Integer
Dim in_coord As POINTAPI
Dim out_coord As POINTAPI
Dim out_coord1 As POINTAPI
Dim out_coord2 As POINTAPI
Dim temp_aid_line As display_line_data
in_coord.X = in_x
in_coord.Y = in_y
Dim is_ty As Integer
 If find_point_from_points(point_no%, m_backup_line_data(0).data.in_point) And point_no% > 0 Then
     read_line = 1
      Exit Function
 End If
'paral_or_verti_ = 0
        If m_backup_line_data(0).data.in_point(0) > 1 And m_backup_line_data(0).data.type = condition Then '普通线
           is_ty = condition
         ElseIf m_backup_line_data(0).data.in_point(0) = 1 And m_backup_line_data(0).data.type = aid_condition Then '辅助线，
           is_ty = aid_condition
        Else
           is_ty = 0
        End If
     If m_backup_line_data(0).aid_line.start_point > 0 Then
      read_line = read_line_(m_object, in_coord, m_backup_line_data(0).aid_line, out_coord, out_coord1, out_coord2, verti_) '垂直
     End If
    If read_line > 0 Then
                   paral_or_verti_ = verti_
                   out_x = out_coord.X
                   out_y = out_coord.Y
                  out1_x = out_coord1.X 'm_backup_line_data(0).aid_line.display_aid_line.end_point_coord(0).X
                  out1_y = out_coord1.Y 'm_backup_line_data(0).aid_line.display_aid_line.end_point_coord(0).Y
                  out2_x = out_coord2.X 'm_backup_line_data(0).aid_line.display_aid_line.end_point_coord(1).X
                  out2_y = out_coord2.Y 'm_backup_line_data(0).aid_line.display_aid_line.end_point_coord(1).Y
    Else
      read_line = read_line_(m_object, in_coord, m_backup_line_data(0).aid_line, out_coord, out_coord1, out_coord2, is_ty)
      If read_line > 0 Then '选中原线
                  out_x = out_coord.X
                   out_y = out_coord.Y
                  out1_x = out_coord1.X 'm_backup_line_data(0).aid_line.display_aid_line.end_point_coord(0).X
                  out1_y = out_coord1.Y 'm_backup_line_data(0).aid_line.display_aid_line.end_point_coord(0).Y
                  out2_x = out_coord2.X 'm_backup_line_data(0).aid_line.display_aid_line.end_point_coord(1).X
                  out2_y = out_coord2.Y 'm_backup_line_data(0).aid_line.display_aid_line.end_point_coord(1).Y
                 paral_or_verti_ = 0
          ' End If
      ElseIf m_backup_line_data(0).aid_line.start_point > 0 Then
        'is_same_POINTAPI(in_coord, m_poi(m_backup_line_data(0).aid_line.start_point).data(0).data0.coordinate) = False Then
      '有辅助线起点，且不是鼠标点
        'If m_backup_line_data(0).aid_line.start_point <> m_backup_line_data(0).end_poi(0) Then '平行线不与原线重合
          read_line = read_line_(m_object, in_coord, m_backup_line_data(0).aid_line, out_coord, out_coord1, out_coord2, paral_) ' 平行
         If read_line > 0 Then
                  out_x = out_coord.X
                  out_y = out_coord.Y
                  out1_x = out_coord1.X 'm_backup_line_data(0).aid_line.display_aid_line.end_point_coord(0).X
                  out1_y = out_coord1.Y 'm_backup_line_data(0).aid_line.display_aid_line.end_point_coord(0).Y
                  out2_x = out_coord2.X 'm_backup_line_data(0).aid_line.display_aid_line.end_point_coord(1).X
                  out2_y = out_coord2.Y 'm_backup_line_data(0).aid_line.display_aid_line.end_point_coord(1).Y
                  paral_or_verti_ = paral_ '读出平行线
         End If
       'End If
      End If
     End If
 End Function
Public Sub set_aid_line(ByVal line_no%, ByVal p1%, ByVal p2%, _
                              ByVal x1 As Long, ByVal y1 As Long, _
                              ByVal x2 As Long, ByVal y2 As Long, _
                              d_cond1_ty As Byte, d_cond1_no%, _
                              d_cond2_ty As Byte, d_cond2_no%, _
                              visible As Byte, color As Byte)
m_line_no = line_no%
m_backup_line_data(0).data.type = aid_condition
m_backup_line_data(0).data.poi(0) = p1%
m_backup_line_data(0).data.poi(1) = p2%
If p1% > 0 Then
   m_backup_line_data(0).display_line0.end_point_coord(0) = m_poi(p1%).data(0).data0.coordinate
Else
   m_backup_line_data(0).display_line0.end_point_coord(0).X = x1
   m_backup_line_data(0).display_line0.end_point_coord(0).Y = y1
End If
If p2% > 0 Then
   m_backup_line_data(0).display_line0.end_point_coord(1) = m_poi(p2%).data(0).data0.coordinate
Else
   m_backup_line_data(0).display_line0.end_point_coord(1).X = x2
   m_backup_line_data(0).display_line0.end_point_coord(1).Y = y2
End If
m_depend_condition(0).ty = d_cond1_ty
m_depend_condition(1).ty = d_cond2_ty
m_depend_condition(0).no = d_cond1_no%
m_depend_condition(1).no = d_cond2_no%
m_backup_line_data(1).data.total_color = color
If p2% > 0 Then
m_backup_line_data(1).data.visible = 1
Call draw_line(1)
End If
End Sub
Public Sub PointCoordinateChange(point_no%)
If is_point_in_line3(point_no%, m_lin(m_line_no).data(0).data0, 0) Then
Call draw_line_(0, 0)
End If
End Sub
Public Sub get_end_point(p1%, p2%, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
p1% = m_backup_line_data(0).data.poi(0)
p2% = m_backup_line_data(0).data.poi(1)
x1 = m_backup_line_data(0).display_line0.end_point_coord(0).X
y1 = m_backup_line_data(0).display_line0.end_point_coord(0).Y
x2 = m_backup_line_data(0).display_line0.end_point_coord(1).X
y2 = m_backup_line_data(0).display_line0.end_point_coord(1).Y
End Sub
Public Function kill_aid_line() As Boolean
If m_backup_line_data(0).data.type = aid_condition Then
    m_lin(m_line_no%).data(0).data0.type = condition
    m_lin(m_line_no%).data(0).data0.total_color = condition_color
    m_lin(m_line_no%).data(0).is_change = True
    kill_aid_line = True
ElseIf m_backup_line_data(0).data.in_point(0) < 2 Then '未画完的线
 Call delete_line(m_line_no)
 'Call Class_Terminate
 kill_aid_line = True
ElseIf m_backup_line_data(0).data.type = condition Then '恢复条件直线
 If m_lin(m_line_no%).data(0).data0.total_color <> condition_color Then
    m_lin(m_line_no%).data(0).data0.total_color = condition_color
     m_lin(m_line_no%).data(0).is_change = True
      Call draw_line_(0, 0)
 kill_aid_line = True
 End If
End If
End Function
Public Sub get_depend_data(d_ty1 As Integer, d_no1%, d_ty2 As Integer, d_no2%)
d_ty1 = m_depend_condition(0).ty
d_ty2 = m_depend_condition(1).ty
d_no1% = m_depend_condition(0).no
d_no2% = m_depend_condition(1).no
End Sub
Public Sub redraw(Optional draw_ty As Integer = 0) '=0画线,=1画线结束
Dim i%
If Backup(draw_ty, 0, 0) Then '备份数据
If m_backup_line_data(0).total_section_no = 1 And m_backup_line_data(1).total_section_no = 1 Then
   If is_different_display_line(m_backup_line_data(0).display_line(1), m_backup_line_data(1).display_line(1)) Then
   Call Drawline(m_object, QBColor(m_backup_line_data(1).display_line(1).color), 0, _
                  m_backup_line_data(1).display_line(1).end_point_coord(0), _
                   m_backup_line_data(1).display_line(1).end_point_coord(1), _
                    m_backup_line_data(1).data.in_point(10), m_backup_line_data(1).data.visible)
   Call Drawline(m_object, QBColor(m_backup_line_data(0).display_line(1).color), 0, _
                  m_backup_line_data(0).display_line(1).end_point_coord(0), _
                   m_backup_line_data(0).display_line(1).end_point_coord(1), _
                    m_backup_line_data(0).data.in_point(10), m_backup_line_data(0).data.visible)
   End If
Else
 For i% = 1 To m_backup_line_data(0).total_section_no
   Call Drawline(m_object, QBColor(m_backup_line_data(1).display_line(i%).color), 0, _
                  m_backup_line_data(1).display_line(i%).end_point_coord(0), _
                   m_backup_line_data(1).display_line(i%).end_point_coord(1), _
                    m_backup_line_data(1).data.in_point(10), m_backup_line_data(1).data.visible)
 Next i%
For i% = 1 To m_backup_line_data(0).total_section_no
   Call Drawline(m_object, QBColor(m_backup_line_data(0).display_line(i%).color), 0, _
                  m_backup_line_data(0).display_line(i%).end_point_coord(0), _
                   m_backup_line_data(0).display_line(i%).end_point_coord(1), _
                    m_backup_line_data(0).data.in_point(10), m_backup_line_data(0).data.visible)
Next i%
End If
End If
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "display_circle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_circle_no As Integer
Private m_backup_data(1) As circle_data0_type
'Private m_backup_center As Integer
'Private m_backup_c_coord  As POINTAPI
'Private m_backup_radii As Long
'Private m_backup_color As Byte
'Private m_backup_visible As Byte
Private m_backup_type As Byte
Private m_object As Object
Private m_picture_range As Long
Private m_picture_orig_x As Long
Private m_picture_orig_y As Long
Private m_picture_top_left_x As Long
Private m_picture_top_left_y As Long
Private is_set_print_data As Boolean
'Private m_data As circle_data_type '1 条件 2 结论 3 辅助
Private father_class As Object
Public Sub picture_orig_x(p_or_x As Long)
  m_picture_orig_x = p_or_x
End Sub
Public Sub picture_orig_y(p_or_y As Long)
  m_picture_orig_y = p_or_y
End Sub
Public Sub picture_range(p_r As Long)
  m_picture_range = p_r
  Call Backup(0, 0)
End Sub
Public Sub picture_top_left_x(top_x As Long)
  m_picture_top_left_x = top_x
End Sub
Public Sub picture_top_left_y(top_y As Long)
  m_picture_top_left_y = top_y
End Sub
Property Get circle_no() As Integer
   circle_no = m_circle_no
End Property
Public Sub set_m_radii(ByVal radii&)
 If m_Circ(m_circle_no%).data(0).data0.radii <> radii Then
    m_Circ(m_circle_no%).data(0).data0.radii = radii&
   Call draw_circle_(0, 0)
 End If
End Sub
Public Sub set_real_radii(ByVal real_radii&)
 If m_Circ(m_circle_no%).data(0).data0.real_radii <> real_radii Then
    m_Circ(m_circle_no%).data(0).data0.real_radii = real_radii&
    '*** Call set_m_radll(ratio_measur * real_radii&)
 End If
End Sub
Public Sub set_change_ratio_measure()
 If m_Circ(m_circle_no%).data(0).data0.real_radii <> 0 Then
   '*** m_Circ(m_circle_no%).data(0).data0.real_radii = real_radii&
   '***  Call set_m_radll(ratio_measur * real_radii&)
 End If
End Sub

Public Property Get color() As Byte
 color = m_Circ(m_circle_no%).data(0).data0.color
End Property
Public Property Get ty() As Byte
 ty = m_Circ(m_circle_no%).data(0).input_type
End Property
Public Sub is_change(ch As Boolean)
 m_Circ(m_circle_no%).data(0).is_change = ch
End Sub
Public Sub set_m_data0(ByVal circ_no%)
m_circle_no = circ_no%
   Call draw_circle_(0, 0)
End Sub
Public Sub set_m_type(ty As Byte, circ_n%)
End Sub

Public Sub set_m_visible(vi As Byte)
 If m_backup_data(1).visible <> vi Then
     m_Circ(m_circle_no%).data(0).data0.visible = vi
       Call draw_circle_(0, 0)
 End If
End Sub
Private Sub draw_circle(Optional color_no As Integer = -1)
'Dim r&
'Dim color As Byte
 'center = m_c_coord
'r& = m_radii
If color_no = -1 Then
If m_backup_data(1).visible = 0 Then
 Exit Sub
End If

If m_backup_data(1).color = 0 Then
If m_backup_type = condition Then
color_no = condition_color
ElseIf m_backup_type = conclusion Then
color_no = conclusion_color
ElseIf m_backup_type = fill Then
color_no = fill_color
End If
Else
color_no = m_backup_data(1).color
End If
End If
m_object.Circle (m_backup_data(1).c_coord.X, m_backup_data(1).c_coord.Y), m_backup_data(1).radii&, QBColor(color_no)
End Sub
Public Sub Backup(Optional X As Long = 0, Optional Y As Long = 0, Optional color_no As Integer = -1)
'
Dim i%
Dim t_p As POINTAPI
t_p.X = X
t_p.Y = Y
If m_picture_range = 0 Then
'If draw_ty = 0 Or draw_ty = 1 Then
If X = 0 And Y = 0 Then
 m_backup_data(1).c_coord = m_Circ(m_circle_no).data(0).data0.c_coord 'fu
 m_backup_data(1).radii = m_Circ(m_circle_no).data(0).data0.radii '
 If m_Circ(m_circle_no).data(0).data0.in_point(0) = 0 And m_Circ(m_circle_no).data(0).data0.color = 0 Then '圆上无点，是辅助圆
  m_backup_data(1).color = fill_color
 Else
  If m_Circ(m_circle_no).data(0).data0.visible = 0 Then
  m_backup_data(1).color = 15
  Else
  m_backup_data(1).color = m_Circ(m_circle_no).data(0).data0.color
  End If
 End If
 m_backup_data(1).visible = m_Circ(m_circle_no).data(0).data0.visible
 m_backup_type = m_Circ(m_circle_no).data(0).input_type
ElseIf m_Circ(m_circle_no).data(0).circle_type = 0 Then
  m_backup_data(1).radii = sqr((m_backup_data(1).c_coord.X - X) ^ 2 + _
                              (m_backup_data(1).c_coord.Y - Y) ^ 2)
  m_Circ(circle_no).data(0).data0.radii = m_backup_data(1).radii
ElseIf m_Circ(m_circle_no).data(0).circle_type = 1 Then '两点圆
  m_backup_data(1).radii = distance_of_two_POINTAPI(m_poi(m_Circ(m_circle_no).data(0).data0.in_point(1)).data(0).data0.coordinate, _
                          t_p) / 2
  m_Circ(circle_no).data(0).data0.radii = m_backup_data(1).radii
  m_Circ(circle_no).data(0).data0.c_coord = _
     mid_POINTAPI(m_poi(m_Circ(m_circle_no).data(0).data0.in_point(1)).data(0).data0.coordinate, _
                        t_p)
ElseIf m_Circ(m_circle_no).data(0).circle_type = 2 Then '三点圆
  m_backup_data(1).radii = circle_radii0(m_poi(m_Circ(m_circle_no).data(0).data0.in_point(1)).data(0).data0.coordinate, _
                                 m_poi(m_Circ(m_circle_no).data(0).data0.in_point(2)).data(0).data0.coordinate, _
                                 t_p, m_backup_data(1).c_coord)
  m_Circ(circle_no).data(0).data0.radii = m_backup_data(1).radii
  m_Circ(circle_no).data(0).data0.c_coord = m_backup_data(1).c_coord
End If
' If draw_ty = 0 Then
  m_backup_data(0) = m_backup_data(1)
  If color_no% <> -1 Then
  m_backup_data(1).color = color_no
  m_backup_data(1).visible = 1
  End If
  
'
'ElseIf draw_ty = 2 Then
' m_backup_data(1) = m_backup_data(0)
' m_Circ(circle_no).data(0).data0 = m_backup_data(0)
'End If
ElseIf is_set_print_data = False Then
   is_set_print_data = True
m_backup_data(1).c_coord.X = (m_Circ(m_circle_no%).data(0).data0.c_coord.X - m_picture_top_left_x) * 2880 / m_picture_range + _
      m_picture_orig_x
m_backup_data(1).c_coord.Y = (m_Circ(m_circle_no%).data(0).data0.c_coord.Y - m_picture_top_left_y) * 2880 / m_picture_range + _
      m_picture_orig_y
m_backup_data(1).radii = m_Circ(m_circle_no%).data(0).data0.radii * 2880 / m_picture_range
End If
End Sub
Public Sub draw_circle_(Optional X As Long = 0, Optional Y As Long = 0, Optional color_no As Integer = -1)
If run_type = 0 Or run_type > 5 Then
   Call draw_circle
   Call Backup(X, Y, color_no)
   Call draw_circle(color_no)
End If
End Sub

Public Sub set_object(ob1 As Object, ob2 As Object)
Set m_object = ob1
Set father_class = ob2
End Sub

Private Sub Class_Initialize()
is_set_print_data = False
End Sub

Private Sub Class_Terminate()
m_Circ(m_circle_no).is_set_data = False
Call draw_circle
End Sub

Public Sub set_m_color(color As Byte)
If m_Circ(m_circle_no%).data(0).data0.color <> color Then
   m_Circ(m_circle_no%).data(0).data0.color = color
   Call draw_circle_(0, 0)
End If
End Sub

Public Sub set_center(ByVal c_center%)
If m_Circ(m_circle_no%).data(0).data0.center = 0 Or _
          m_Circ(m_circle_no%).data(0).data0.c_coord.X <> m_poi(m_Circ(m_circle_no%).data(0).data0.center).data(0).data0.coordinate.X Or _
          m_Circ(m_circle_no%).data(0).data0.c_coord.Y <> m_poi(m_Circ(m_circle_no%).data(0).data0.center).data(0).data0.coordinate.Y Then
m_Circ(m_circle_no%).data(0).data0.center = c_center%
m_Circ(m_circle_no%).data(0).data0.c_coord = m_poi(c_center).data(0).data0.coordinate
m_Circ(m_circle_no%).data(0).data0.radii = abs_POINTAPI(minus_POINTAPI(m_Circ(m_circle_no%).data(0).data0.c_coord, _
                       m_poi(m_Circ(m_circle_no%).data(0).data0.in_point(1)).data(0).data0.coordinate))
Call draw_circle_(0, 0)
End If
End Sub
Public Sub print_(ob As Object)
  Set m_object = ob
   Call Backup(0, 0)
    Call draw_circle
End Sub
Public Function is_char_non_interset_circle( _
        char_left_top_x As Long, char_left_top_y As Long) As Boolean
Dim d(3) As Integer
Dim tp(3) As POINTAPI
Dim i%
tp(2).X = char_left_top_x
tp(2).Y = char_left_top_y
tp(3).X = tp(2).X + 110
tp(3).Y = tp(2).Y + 140
tp(0).X = tp(2).X
tp(0).Y = tp(3).Y
tp(1).Y = tp(2).Y
tp(1).X = tp(3).X
For i% = 0 To 3
d(i%) = point_in_circle(tp(i%), m_backup_data(1).c_coord, m_backup_data(1).radii)
Next i%
If d(0) = 1 Then
 If d(1) = 1 And d(2) = 1 And d(3) = 1 Then
   is_char_non_interset_circle = True
 Else
   is_char_non_interset_circle = False
 End If
ElseIf d(0) = 0 Then
is_char_non_interset_circle = False
Else
 If d(1) = -1 And d(2) = -1 And d(3) = -1 Then
  is_char_non_interset_circle = True
 Else
  is_char_non_interset_circle = False
 End If
End If
End Function

Public Function read_m_circle(X As Long, Y As Long) As Integer
Dim coord As POINTAPI
Dim r As Long
    coord.X = X
    coord.Y = Y
 r = abs_POINTAPI(minus_POINTAPI(m_Circ(m_circle_no%).data(0).data0.c_coord, coord))
 r = Abs(r - m_Circ(m_circle_no%).data(0).data0.radii)
 If r < 5 Then
     read_m_circle = m_circle_no
 End If
End Function

Public Sub redraw(Optional is_recover As Boolean = False)
If is_recover Then
    Call draw_circle_(0, 0, CInt(m_backup_data(0).color))
Else
If m_Circ(m_circle_no%).data(0).input_type = condition Then
 Call set_m_color(condition_color)
ElseIf m_Circ(m_circle_no%).data(0).input_type = conclusion Then
 Call set_m_color(conclusion_color)
ElseIf m_Circ(m_circle_no%).data(0).input_type = aid_condition Then
 Call Class_Terminate
End If
End If
End Sub
Public Sub PointCoordinateChange(ByVal point_no%)
Dim tcoord As POINTAPI
Dim i%
Dim p%
If m_Circ(m_circle_no%).data(0).circle_type = 1 Then
   If m_Circ(m_circle_no).data(0).data0.center = point_no% Or _
             m_Circ(m_circle_no).data(0).data0.in_point(0) = point_no% Then
       m_Circ(m_circle_no).data(0).data0.c_coord = m_poi(point_no%).data(0).data0.coordinate
   ElseIf m_Circ(m_circle_no).data(0).radii_depend_poi(0) > 0 And _
            m_Circ(m_circle_no).data(0).radii_depend_poi(1) > 0 Then
      If m_Circ(m_circle_no).data(0).radii_depend_poi(0) = point_no% Or _
             m_Circ(m_circle_no).data(0).radii_depend_poi(1) = point_no% Then
             ' Exit Sub
      Else
         For i% = 1 To m_Circ(m_circle_no).data(0).data0.in_point(0)
           If point_no% <= m_Circ(m_circle_no).data(0).data0.in_point(i%) Then
               p% = m_Circ(m_circle_no).data(0).data0.in_point(i%)
            If m_poi(p%).data(0).parent.element(0).ty = circle_ And _
                m_poi(p%).data(0).parent.element(0).no = m_circle_no And _
                 m_poi(p%).data(0).parent.element(1).no = 0 Then
              If read_circle0(m_Circ(m_circle_no).data(0).data0, _
                       m_poi(p%).data(0).data0.coordinate, tcoord) Then
                 If p% > point_no% Then
                  Call set_point_coordinate(p%, tcoord, True)
                 Else
                  m_poi(p%).data(0).data0.coordinate = tcoord
                 End If
              End If
            End If
          End If
         Next i%
             Exit Sub
       End If
   ElseIf m_Circ(m_circle_no).data(0).data0.in_point(1) <> point_no% Then
      Exit Sub
   End If
Else
   If m_Circ(m_circle_no).data(0).data0.in_point(1) <> point_no% And _
      m_Circ(m_circle_no).data(0).data0.in_point(2) <> point_no% And _
      m_Circ(m_circle_no).data(0).data0.in_point(3) <> point_no% Then
       Exit Sub
   End If
End If
      m_Circ(m_circle_no).data(0).is_change = True
      m_Circ(m_circle_no).data(0).data0.radii = _
         m_circle_radii(m_Circ(m_circle_no).data(0))
Call draw_circle_(0, 0)
End Sub


Public Function top_left_x() As Long
    top_left_x = _
       m_Circ(m_circle_no).data(0).data0.c_coord.X - CLng(m_Circ(m_circle_no).data(0).data0.radii)
End Function
Public Function top_left_y() As Long
    top_left_y = _
       m_Circ(m_circle_no).data(0).data0.c_coord.Y - CLng(m_Circ(m_circle_no).data(0).data0.radii)
End Function
       
Public Function bottom_right_x() As Long
    bottom_right_x = _
       m_Circ(m_circle_no).data(0).data0.c_coord.X + CLng(m_Circ(m_circle_no).data(0).data0.radii)
End Function
Public Function bottom_right_y() As Long
    bottom_right_y = _
       m_Circ(m_circle_no).data(0).data0.c_coord.Y + CLng(m_Circ(m_circle_no).data(0).data0.radii)
End Function
Public Sub set_radii_no(radii_no%)
m_Circ(m_circle_no).data(0).radii_no = radii_no%
End Sub

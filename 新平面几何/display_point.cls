VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "display_point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Public data As point_data0_type
Public m_point_no As Integer
'***************************
Private m_backup_name As String '* 1
Private m_backup_coordinate  As POINTAPI
Private m_backup_color As Byte
Private m_backup_visible As Byte
Private m_object As Object
Private m_picture_range As Long
Private m_picture_orig_x As Long
Private m_picture_orig_y As Long
Private m_picture_top_left_x As Long
Private m_picture_top_left_y As Long
Private is_set_print_data As Boolean
Private m_point_name_position As Byte
Private is_set_name_position As Boolean
Private father_class As Object '容器对象
Public Sub picture_orig_x(p_or_x&)
  m_picture_orig_x = p_or_x
End Sub
Public Sub picture_orig_y(p_or_y&)
  m_picture_orig_y = p_or_y
End Sub
Public Sub picture_range(p_r As Long)
  m_picture_range = p_r
  Call Backup
End Sub
Public Sub picture_top_left_x(top_x As Long)
  m_picture_top_left_x = top_x
End Sub
Public Sub picture_top_left_y(top_y As Long)
  m_picture_top_left_y = top_y
End Sub
'Private m_is_display As Boolean
Public Property Get name() As String
 name = m_poi(m_point_no).data(0).data0.name
End Property
Public Sub Set_m_name(name$)
 If ((name$ >= "A" And name$ <= "Z") Or name = "") And m_backup_name <> name$ Then
  m_poi(m_point_no).data(0).data0.name = name
  Call draw_m_poi_
 End If
End Sub

'Public Property Get point_no() As Integer
' point_no = m_point_no
'End Property
Public Sub Set_m_coordinate(X As Long, Y As Long)
   If X <> m_backup_coordinate.X Or _
        Y <> m_backup_coordinate.Y Then
     m_poi(m_point_no).data(0).data0.coordinate.X = X
     m_poi(m_point_no).data(0).data0.coordinate.Y = Y
    Call draw_m_poi_
    Call PointCoordinateChange(m_point_no)  '发送点改变的事件
    If C_display_picture.m_draw_statue = op_change_picture Then
        m_poi(m_point_no).data(0).is_change = True
    End If
   End If
End Sub
Public Property Get color() As Byte
 color = m_poi(m_point_no).data(0).data0.color
End Property
Public Sub set_m_color(co As Byte)
 If m_backup_color <> co Then
     m_poi(m_point_no).data(0).data0.color = co
      Call draw_m_poi_
 End If
End Sub
Public Sub set_m_color_(co As Byte) '用于flash
  Call draw_m_poi
       m_backup_color = co
  Call draw_m_poi
End Sub
Public Property Get visible() As Byte
 visible = m_poi(m_point_no).data(0).data0.visible
End Property
Public Sub set_m_visible(vi As Byte)
  If m_backup_visible <> vi Then
   m_poi(m_point_no).data(0).data0.visible = vi
    Call draw_m_poi_
  End If
End Sub
Private Sub Class_Initialize()
m_picture_range = 0
is_set_print_data = False
is_set_name_position = False
Call Backup
End Sub
Private Sub draw_m_poi()      ', col%,display_or_delete as boolean )  '???????
'draw_ty=0 draw, =1 remove
Dim dis_or_de As Byte
Dim t_name As String '* 1
Dim t_coordinate  As POINTAPI
Dim t_color As Byte
Dim t_fill As Integer
Dim t_visible As Byte
If ((m_backup_name >= "A" And m_backup_name <= "Z") Or m_backup_name = "") And m_backup_visible = 1 Then
 t_name = m_backup_name
 t_coordinate = m_backup_coordinate
 t_color = m_backup_color
 If t_coordinate.X <> 10000 And t_coordinate.Y <> 10000 Then
  If m_picture_range = 0 Then
    Call m_BPset(m_object, t_coordinate, t_name, t_color)
  Else
 t_fill = m_object.fillstyle
 m_object.FillColor = QBColor(0)
 m_object.fillstyle = 0
 m_object.Circle (m_backup_coordinate.X, m_backup_coordinate.Y), 30, QBColor(9)
 If is_set_name_position = False Then
   Call chose_point_name_position
    is_set_name_position = True
 End If
    Call print_name_
 m_object.fillstyle = t_fill
 End If
 End If
 End If
End Sub
'Public Sub set_m_point_no(ByVal point_no%)
'm_point_no = point_no%
'End Sub
Public Sub Backup()
If m_picture_range = 0 Then '显示
m_backup_name = m_poi(m_point_no).data(0).data0.name
m_backup_coordinate = m_poi(m_point_no).data(0).data0.coordinate
m_backup_color = m_poi(m_point_no).data(0).data0.color
m_backup_visible = m_poi(m_point_no).data(0).data0.visible
'm_backup_name_icon = m_name_icon
ElseIf is_set_print_data = False Then
   is_set_print_data = True
m_backup_coordinate.X = (m_poi(m_point_no).data(0).data0.coordinate.X - m_picture_top_left_x) * 2880 / _
         m_picture_range + m_picture_orig_x
m_backup_coordinate.Y = (m_poi(m_point_no).data(0).data0.coordinate.Y - m_picture_top_left_y) * 2880 / _
         m_picture_range + m_picture_orig_y
End If
End Sub

Private Sub Class_Terminate()
m_poi(m_point_no).is_set_data = False
Call draw_m_poi
End Sub
Public Sub set_object(ob As Object, ob1 As Object)
Set m_object = ob
Set father_class = ob1
End Sub
Public Sub print_(ob As Object)
 Set m_object = ob
  Call Backup
   Call draw_m_poi
End Sub

Private Sub chose_point_name_position()
Dim left_top As POINTAPI
Dim right_bottum As POINTAPI
left_top.X = m_backup_coordinate.X + 20 '右下
left_top.Y = m_backup_coordinate.Y + 20
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then '
   m_point_name_position = 0
    Exit Sub
End If
left_top.X = m_backup_coordinate.X + 20 '右上
left_top.Y = m_backup_coordinate.Y - 240 ' 180
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 1
    Exit Sub
End If
left_top.X = m_backup_coordinate.X - 140 '
left_top.Y = m_backup_coordinate.Y - 240 ' 180
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 2
    Exit Sub
End If
left_top.X = m_backup_coordinate.X - 140
left_top.Y = m_backup_coordinate.Y + 20
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 3
    Exit Sub
End If
left_top.X = m_backup_coordinate.X + 35
left_top.Y = m_backup_coordinate.Y - 150 '90
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 4
    Exit Sub
End If
left_top.X = m_backup_coordinate.X - 60
left_top.Y = m_backup_coordinate.Y - 260 ' 200
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 5
    Exit Sub
End If
left_top.X = m_backup_coordinate.X + 180
left_top.Y = m_backup_coordinate.Y - 140 '80
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 6
    Exit Sub
End If
left_top.X = m_backup_coordinate.X - 60
left_top.Y = m_backup_coordinate.Y + 35
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 7
    Exit Sub
End If
left_top.X = m_backup_coordinate.X + 40
left_top.Y = m_backup_coordinate.Y + 40
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 8
    Exit Sub
End If
left_top.X = m_backup_coordinate.X + 40
left_top.Y = m_backup_coordinate.Y - 260 ' 200
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 9
    Exit Sub
End If
left_top.X = m_backup_coordinate.X - 160
left_top.Y = m_backup_coordinate.Y - 260 '200
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 10
    Exit Sub
End If
left_top.X = m_backup_coordinate.X - 160
left_top.Y = m_backup_coordinate.Y + 40
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 11
    Exit Sub
End If
left_top.X = m_backup_coordinate.X + 55
left_top.Y = m_backup_coordinate.Y - 150 '90
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 12
    Exit Sub
End If
left_top.X = m_backup_coordinate.X - 60
left_top.Y = m_backup_coordinate.Y - 280 ' 220
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 13
    Exit Sub
End If
left_top.X = m_backup_coordinate.X - 200
left_top.Y = m_backup_coordinate.Y - 140 ' 80
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 7
    Exit Sub
End If
left_top.X = m_backup_coordinate.X - 60
left_top.Y = m_backup_coordinate.Y + 55
If father_class.chose_point_name_position(left_top.X, left_top.Y) Then
   m_point_name_position = 7
    Exit Sub
End If
End Sub
Public Sub print_name_()
Select Case m_point_name_position
Case 0
m_object.CurrentX = m_backup_coordinate.X + 20
m_object.CurrentY = m_backup_coordinate.Y + 20
Case 8
m_object.CurrentX = m_backup_coordinate.X + 40
m_object.CurrentY = m_backup_coordinate.Y + 40
Case 1
m_object.CurrentX = m_backup_coordinate.X + 20
m_object.CurrentY = m_backup_coordinate.Y - 240  '180
Case 9
m_object.CurrentX = m_backup_coordinate.X + 40
m_object.CurrentY = m_backup_coordinate.Y - 260  '200
Case 2
m_object.CurrentX = m_backup_coordinate.X - 140
m_object.CurrentY = m_backup_coordinate.Y - 240  '180
Case 10
m_object.CurrentX = m_backup_coordinate.X - 160
m_object.CurrentY = m_backup_coordinate.Y - 260  '200
Case 3
m_object.CurrentX = m_backup_coordinate.X - 140
m_object.CurrentY = m_backup_coordinate.Y + 20
Case 11
m_object.CurrentX = m_backup_coordinate.X - 160
m_object.CurrentY = m_backup_coordinate.Y + 40
Case 4
m_object.CurrentX = m_backup_coordinate.X + 35
m_object.CurrentY = m_backup_coordinate.Y - 150 '90
Case 12
m_object.CurrentX = m_backup_coordinate.X + 55
m_object.CurrentY = m_backup_coordinate.Y - 150  '90
Case 5
m_object.CurrentX = m_backup_coordinate.X - 60
m_object.CurrentY = m_backup_coordinate.Y - 260  '200
Case 13
m_object.CurrentX = m_backup_coordinate.X - 60
m_object.CurrentY = m_backup_coordinate.Y - 220
Case 6
m_object.CurrentX = m_backup_coordinate.X - 180
m_object.CurrentY = m_backup_coordinate.Y - 140  '80
Case 14
m_object.CurrentX = m_backup_coordinate.X - 200
m_object.CurrentY = m_backup_coordinate.Y - 140  '80
Case 7
m_object.CurrentX = m_backup_coordinate.X - 60
m_object.CurrentY = m_backup_coordinate.Y + 35
Case 15
m_object.CurrentX = m_backup_coordinate.X - 60
m_object.CurrentY = m_backup_coordinate.Y + 55
End Select
 m_object.Print m_backup_name
End Sub
Public Sub set_father_class(ob As Object)
 Set father_class = ob
End Sub
Public Sub CircleChange(ByVal circle_no%)
Dim tcoord As POINTAPI
If m_poi(m_point_no).data(0).parent.element(0).ty = circle_ And _
    m_poi(m_point_no).data(0).parent.element(0).no = circle_no And _
     m_poi(m_point_no).data(0).parent.element(1).no = 0 Then
     If read_circle(m_point_no, m_poi(m_point_no).data(0).data0.coordinate, _
              tcoord) Then
        Call set_point_coordinate(m_point_no, tcoord, True)
     End If
End If
End Sub

Public Sub draw_m_poi_()
If run_type = 0 Or run_type > 5 Then
 Call draw_m_poi
 Call Backup
 Call draw_m_poi
End If
End Sub

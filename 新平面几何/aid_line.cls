VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "aid_line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type aid_line_data_type
  head_point As Integer
  point_coord(2) As POINTAPI 'coord(0)起点, coord(1)盘终点, coord(2)，方向
  director_type As Integer '(|p0y-p2y|<3,p0x<p2x =0,p0x>p2x=1),p0y<p2y=2,p0y>p2y =3
  visible As Boolean
End Type
Dim data As aid_line_data_type
Dim m_object As Object
Dim father_class As Object
Public Function set_aid_line(head_point0 As Integer, end_point_X As Long, end_point_Y As Long) As Boolean
If head_point = 0 Then
    set_aid_line = False
Else
   data.head_point = head_point0
    data.point_coord(0).X = m_poi(head_point).data(0).data0.coordinate.X
      data.point_coord(0).Y = m_poi(head_point).data(0).data0.coordinate.Y
        data.point_coord(2).X = end_point_X
          data.point_coord(2).Y = end_point_Y
           data.visible = False
            set_aid_line = True
End If
End Function
Public Sub set_object(ob As Object, ob1 As Object)
Set m_object = ob
Set father_class = ob1
End Sub

Public Sub add_point_at_aid_line(in_point_X As Long, in_point_Y As Long)
Dim p_coord As POINTAPI
Dim pointapi0 As POINTAPI
Dim dis As Integer
p_coord.X = in_point_X
p_coord.Y = in_point_Y
 If distance_point_to_line(p_coord, data.point_coord(0), paral_, data.point_coord(0), data.point_coord(2), dis, data.point_coord(1)) Then
    If Abs(dis) < 5 Then
        data.point_coord = point0
          data.visible = True
          Call draw_aid_line
    End If
 ElseIf data.visible = True Then
         data.visible = False
          Call draw_aid_line
 End If
End Sub
Public Function is_me(head_poin0 As Integer, end_point_coord_x, end_point_coord_y) As Boolean
Dim end_p As POINTAPI
 end_p.X = end_point_coord_x
 end_p.Y = end_point_coord_y
 If data.head_point = head_point0 Then
   If is_point_on_line(end_p, point_coord(0), point_coord(2)) > 0 Then
      is_me = True
   End If
 End If
End Function

Private Sub Class_Initialize()
  data.end_poi(0) = 0
  data.end_poi(1) = 0
  data.point_coord(0).X = 10000
  data.point_coord(0).Y = 10000
  data.point_coord(1).X = 10000
  data.point_coord(1).Y = 10000
  data.point_coord(2).X = 10000
  data.point_coord(2).Y = 10000
  data.visible = False
End Sub


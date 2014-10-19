Attribute VB_Name = "change_picture_module"
Option Explicit
Public Sub change_polygon(mo_d As POINTAPI, A!, similar_ratio!)
Dim i%
Dim X&, Y&, d_no%
If is_first_move Then
 For i% = 1 To Polygon_for_change.p(0).total_v - 1
' Call C_display_picture.redraw_point(Polygon_for_change.p(0).v(i%))
 Call redraw_red_line(line_number0(Polygon_for_change.p(0).v(i%), _
     Polygon_for_change.p(0).v(i% - 1), 0, 0))
 Next i%
 'Call C_display_picture.redraw_point(Polygon_for_change.p(0).v(0))
 If Polygon_for_change.p(0).total_v > 2 Then
 Call redraw_red_line(line_number0(Polygon_for_change.p(0).v(0), _
     Polygon_for_change.p(0).v(Polygon_for_change.p(0).total_v - 1), 0, 0))
     is_first_move = False
 End If

Else ' If
Call draw_polygon(Polygon_for_change.p(0), Polygon_for_change.direction) '消图
End If
Polygon_for_change.similar_ratio = similar_ratio!
Polygon_for_change.move = add_pointapi(Polygon_for_change.move, mo_d)
If A <> 0 Then
Polygon_for_change.rote_angle = Polygon_for_change.rote_angle + A
If Polygon_for_change.rote_angle > 2 * PI Then
 Polygon_for_change.rote_angle = Polygon_for_change.rote_angle - 2 * PI
ElseIf Polygon_for_change.rote_angle < -2 * PI Then
 Polygon_for_change.rote_angle = Polygon_for_change.rote_angle + 2 * PI
End If
End If
Polygon_for_change.p(0).coord_center = add_pointapi(Polygon_for_change.p(0).center, Polygon_for_change.move)
'Polygon_for_change.p(0).coord_center.X = _
   Polygon_for_change.p(0).center.X + _
     Polygon_for_change.move.X
'Polygon_for_change.p(0).coord_center.Y = _
  Polygon_for_change.p(0).center.Y + _
    Polygon_for_change.move.Y
If Polygon_for_change.direction = 1 Then
   d_no% = 0
Else
   d_no% = 1
End If
For i% = 0 To Polygon_for_change.p(0).total_v - 1
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
          Polygon_for_change.p(0).coord(i%).X = X&
          Polygon_for_change.p(0).coord(i%).Y = Y&
Next i%
Call draw_polygon(Polygon_for_change.p(0), Polygon_for_change.direction)
End Sub
Public Sub change_line(mo_x&, _
 mo_y&, A!, similar_ratio!)
Dim i%, data_n%
Dim X&, Y&
If is_first_move Then
  'Call C_display_picture.redraw_point(line_for_change.lin(0).poi(0))
  'Call C_display_picture.redraw_point(line_for_change.lin(0).poi(1))
  Call redraw_red_line(line_number0(line_for_change.lin(0).poi(0), _
     line_for_change.lin(0).poi(1), 0, 0))
      is_first_move = False
Else ' If
 Call draw_change_line(5)
End If
line_for_change.similar_ratio = similar_ratio!
line_for_change.move.X = line_for_change.move.X + mo_x&
line_for_change.move.Y = line_for_change.move.Y + mo_y&
If A <> 0 Then
line_for_change.rote_angle = line_for_change.rote_angle + A
If line_for_change.rote_angle > 2 * PI Then
 line_for_change.rote_angle = line_for_change.rote_angle - 2 * PI
ElseIf line_for_change.rote_angle < -2 * PI Then
 line_for_change.rote_angle = line_for_change.rote_angle + 2 * PI
End If
End If
line_for_change.lin(0).center(1).X = _
 line_for_change.lin(0).center(0).X + _
     line_for_change.move.X
line_for_change.lin(0).center(1).Y = _
  line_for_change.lin(0).center(0).Y + _
    line_for_change.move.Y
For i% = 0 To 1
If line_for_change.direction = 1 Then
 data_n% = 0
Else
 data_n% = 1
End If
line_for_change.lin(0).coord(i%).X = line_for_change.lin(0).center(1).X + _
     ((m_poi(line_for_change.lin(0).poi(i%)).data(0).data0.coordinate.X - _
       line_for_change.lin(0).center(0).X) * _
       Cos(line_for_change.rote_angle) - _
        (m_poi(line_for_change.lin(0).poi(i%)).data(0).data0.coordinate.Y - _
         line_for_change.lin(0).center(0).Y) * _
          Sin(line_for_change.rote_angle)) * _
           line_for_change.similar_ratio
line_for_change.lin(0).coord(i%).Y = line_for_change.lin(0).center(1).Y + _
     ((m_poi(line_for_change.lin(0).poi(i%)).data(0).data0.coordinate.X - _
        line_for_change.lin(0).center(0).X) * _
      Sin(line_for_change.rote_angle) + _
        (m_poi(line_for_change.lin(0).poi(i%)).data(0).data0.coordinate.Y - _
          line_for_change.lin(0).center(0).Y) * _
           Cos(line_for_change.rote_angle)) * _
          line_for_change.similar_ratio
Next i%
For i% = 1 To line_for_change.lin(0).in_point(0)
 line_for_change.lin(0).in_poi_coord(i%).X = line_for_change.lin(0).center(1).X + _
     ((m_poi(line_for_change.lin(0).in_point(i%)).data(0).data0.coordinate.X - _
        line_for_change.lin(0).center(0).X) * _
          Cos(line_for_change.rote_angle) - _
      (m_poi(line_for_change.lin(0).in_point(i%)).data(0).data0.coordinate.Y - _
        line_for_change.lin(0).center(0).Y) * _
          Sin(line_for_change.rote_angle)) * _
           line_for_change.similar_ratio
 line_for_change.lin(0).in_poi_coord(i%).Y = line_for_change.lin(0).center(1).Y + _
     ((m_poi(line_for_change.lin(0).in_point(i%)).data(0).data0.coordinate.X - _
        line_for_change.lin(0).center(0).X) * _
          Sin(line_for_change.rote_angle) + _
      (m_poi(line_for_change.lin(0).in_point(i%)).data(0).data0.coordinate.Y - _
        line_for_change.lin(0).center(0).Y) * _
          Cos(line_for_change.rote_angle)) * _
            line_for_change.similar_ratio
Next i%
Call draw_change_line(5)
End Sub
Public Sub change_polygon1(mo_x&, _
 mo_y&)
Dim i%
Dim X&, Y&
Dim t!, r&
Call draw_polygon(Polygon_for_change.p(0), Polygon_for_change.direction)
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
Call draw_polygon(Polygon_for_change.p(0), Polygon_for_change.direction)

End Sub
Public Sub change_polygon2()
Dim i%
Dim X&, Y&
Dim t!, r&
Call draw_polygon(Polygon_for_change.p(0), Polygon_for_change.direction)
For i% = 0 To Polygon_for_change.p(0).total_v - 1
 X& = 2 * center_p.X - Polygon_for_change.p(0).coord(i%).X
 Y& = 2 * center_p.Y - Polygon_for_change.p(0).coord(i%).Y

          Polygon_for_change.p(0).coord(i%).X = X&
          Polygon_for_change.p(0).coord(i%).Y = Y&
Next i%
Call draw_polygon(Polygon_for_change.p(0), Polygon_for_change.direction)

End Sub
Public Sub change_circle(mo_x&, mo_y&, rote_angle!, similar_ratio!)
Dim i%, data_no%
Dim X&, Y&
If is_first_move Then
  Call draw_change_circle(1)
  is_first_move = False
Else ' If
  Call draw_change_circle_
End If
Circle_for_change.move.X = Circle_for_change.move.X + mo_x&
 Circle_for_change.move.Y = Circle_for_change.move.Y + mo_y&
  Circle_for_change.c_coord.X = Circle_for_change.move.X + m_Circ(Circle_for_change.c).data(0).data0.c_coord.X
  Circle_for_change.c_coord.Y = Circle_for_change.move.Y + m_Circ(Circle_for_change.c).data(0).data0.c_coord.Y
If Circle_for_change.direction = 1 Then
 data_no% = 0
Else
 data_no% = 1
End If
For i% = 1 To m_Circ(Circle_for_change.c).data(0).data0.in_point(0)
 Circle_for_change.poi_coordinate(i%).X = Circle_for_change.c_coord.X + _
     ((m_poi(m_Circ(Circle_for_change.c).data(0).data0.in_point(i%)).data(0).data0.coordinate.X - _
        m_Circ(Circle_for_change.c).data(0).data0.c_coord.X) * _
          Cos(Circle_for_change.rote_angle) - _
      (m_poi(m_Circ(Circle_for_change.c).data(0).data0.in_point(i%)).data(0).data0.coordinate.Y - _
        m_Circ(Circle_for_change.c).data(0).data0.c_coord.Y) * _
          Sin(Circle_for_change.rote_angle)) * _
           Circle_for_change.similar_ratio
 Circle_for_change.poi_coordinate(i%).Y = Circle_for_change.c_coord.Y + _
     ((m_poi(m_Circ(Circle_for_change.c).data(0).data0.in_point(i%)).data(0).data0.coordinate.X - _
        m_Circ(Circle_for_change.c).data(0).data0.c_coord.X) * _
          Sin(Circle_for_change.rote_angle) + _
      (m_poi(m_Circ(Circle_for_change.c).data(0).data0.in_point(i%)).data(0).data0.coordinate.Y - _
        m_Circ(Circle_for_change.c).data(0).data0.c_coord.Y) * _
          Cos(Circle_for_change.rote_angle)) * _
            Circle_for_change.similar_ratio
Next i%
 Circle_for_change.radii = m_Circ(Circle_for_change.c).data(0).data0.radii * Circle_for_change.similar_ratio
  Call draw_change_circle_
End Sub
Public Sub oxis_symmetry_polygon(p As polygon, oxis_p1 As POINTAPI, _
                    oxis_p2 As POINTAPI)
Dim c1&, c2&, A1&, b1&, d&
Dim i%
If Abs(oxis_p1.X - oxis_p2.X) + Abs(oxis_p1.Y - oxis_p2.Y) < 8 Then
   Exit Sub
End If
For i% = 0 To p.total_v - 1
'  poi(p.v(i%)).data(1).data(0).data0.coordinate.X = _
        2 * p.center.X - poi(p.v(i%)).data(0).data0.coordinate.X
'  poi(p.v(i%)).data(1).data(0).data0.coordinate.Y = _
               poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.Y
Next i%
A1& = oxis_p1.X - oxis_p2.X
b1& = oxis_p1.Y - oxis_p2.Y
'A2& = b1&
'b2& = -A1&
d& = A1& * A1& + b1& * b1&
If d& > 0 Then
For i% = 0 To p.total_v - 1
c1& = A1& * m_poi(p.v(i%)).data(0).data0.coordinate.X + b1& * m_poi(p.v(i%)).data(0).data0.coordinate.Y
c2& = b1& * (2 * oxis_p1.X - m_poi(p.v(i%)).data(0).data0.coordinate.X) - _
 A1& * (2 * oxis_p1.Y - m_poi(p.v(i%)).data(0).data0.coordinate.Y)
 p.coord(i%).X = (b1& * c2& + A1& * c1&) / d&
 p.coord(i%).Y = (-A1& * c2& + b1& * c1&) / d&
Next i%
End If
End Sub
Public Sub turn_over_polygon(oxis_p1 As POINTAPI, oxis_p2 As POINTAPI)
Dim c1&, c2&, A1&, b1&, d&
Dim i%
If Abs(oxis_p1.X - oxis_p2.X) + Abs(oxis_p1.Y - oxis_p2.Y) < 8 Then
   Exit Sub
End If
For i% = 0 To Polygon_for_change.p(0).total_v - 1
'  poi(Polygon_for_change.p(0).v(i%)).data(1).data(0).data0.coordinate.X = _
        2 * Polygon_for_change.p(0).center.X - poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.X
'  poi(Polygon_for_change.p(0).v(i%)).data(1).data(0).data0.coordinate.Y = _
               poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.Y
Next i%
For i% = 0 To Polygon_for_change.p(0).total_v - 1
 Polygon_for_change.p(0).coord(i%) = oxis_symmetry_point( _
        Polygon_for_change.p(0).coord(i%), oxis_p1, oxis_p2)
Next i%
 Polygon_for_change.p(0).coord_center = oxis_symmetry_point( _
        Polygon_for_change.p(0).coord_center, oxis_p1, oxis_p2)
  Polygon_for_change.move.X = Polygon_for_change.p(0).coord_center.X - Polygon_for_change.p(0).center.X
  Polygon_for_change.move.Y = Polygon_for_change.p(0).coord_center.Y - Polygon_for_change.p(0).center.Y
Polygon_for_change.direction = -Polygon_for_change.direction
If Polygon_for_change.direction = 1 Then
   Polygon_for_change.rote_angle = angle_value_from_two_line(Polygon_for_change.p(0).coord(1), _
       Polygon_for_change.p(0).coord(0), m_poi(Polygon_for_change.p(0).v(1)).data(0).data0.coordinate, _
          m_poi(Polygon_for_change.p(0).v(0)).data(0).data0.coordinate)
Else
'   Polygon_for_change.rote_angle = angle_value_from_two_line(Polygon_for_change.p(0).coord(1), _
       Polygon_for_change.p(0).coord(0), m_m_poi(Polygon_for_change.p(0).v(1)).data(1).data(0).data0.coordinate, _
          m_m_poi(Polygon_for_change.p(0).v(0)).data(1).data(0).data0.coordinate)
End If
is_first_move = False
End Sub
Public Sub turn_over_circle(oxis_p1 As POINTAPI, oxis_p2 As POINTAPI)
Dim c1&, c2&, A1&, b1&, d&
Dim i%
If Abs(oxis_p1.X - oxis_p2.X) + Abs(oxis_p1.Y - oxis_p2.Y) < 8 Then
   Exit Sub
End If
For i% = 1 To m_Circ(Circle_for_change.c).data(0).data0.in_point(0)
 Circle_for_change.poi_coordinate(i%) = oxis_symmetry_point( _
        Circle_for_change.poi_coordinate(i%), oxis_p1, oxis_p2)
Next i%
 Circle_for_change.c_coord = oxis_symmetry_point( _
       Circle_for_change.c_coord, oxis_p1, oxis_p2)
  Circle_for_change.move.X = Circle_for_change.c_coord.X - m_Circ(Circle_for_change.c).data(0).data0.c_coord.X
  Circle_for_change.move.Y = Circle_for_change.c_coord.Y - m_Circ(Circle_for_change.c).data(0).data0.c_coord.Y
 Circle_for_change.direction = -Circle_for_change.direction
If Circle_for_change.direction = 1 Then
   Circle_for_change.rote_angle = angle_value_from_two_line(Circle_for_change.poi_coordinate(1), _
       Circle_for_change.c_coord, m_poi(m_Circ(Circle_for_change.c).data(0).data0.in_point(1)).data(0).data0.coordinate, _
          m_Circ(Circle_for_change.c).data(0).data0.c_coord)
Else
 '  Circle_for_change.rote_angle = angle_value_from_two_line(Circle_for_change.poi_coordinate(1), _
       Circle_for_change.c_coord, m_poi(m_circ(Circle_for_change.c).data(0).data0.in_point(1)).data(1).data(0).data0.coordinate, _
          m_circ(Circle_for_change.c).data(0).data0.c_coord)
End If
is_first_move = False
End Sub
Public Sub oxis_symmetry_line(oxis_p1 As POINTAPI, _
                    oxis_p2 As POINTAPI)
Dim i%
If Abs(oxis_p1.X - oxis_p2.X) + Abs(oxis_p1.Y - oxis_p2.Y) < 8 Then
   Exit Sub
End If
For i% = 1 To line_for_change.lin(0).in_point(0)
 line_for_change.lin(0).in_poi_coord(i%) = oxis_symmetry_point( _
        m_poi(line_for_change.lin(0).in_point(i%)).data(0).data0.coordinate, oxis_p1, oxis_p2)
Next i%
 line_for_change.lin(0).center(1) = oxis_symmetry_point( _
        line_for_change.lin(0).center(0), oxis_p1, oxis_p2)
line_for_change.lin(0).coord(0) = line_for_change.lin(0).in_poi_coord(1)
line_for_change.lin(0).coord(1) = line_for_change.lin(0).in_poi_coord(line_for_change.lin(0).in_point(0))
End Sub
Public Sub turn_over_line(oxis_p1 As POINTAPI, _
                            oxis_p2 As POINTAPI)
Dim i%
Dim na!
Dim tp(1) As POINTAPI
If Abs(oxis_p1.X - oxis_p2.X) + Abs(oxis_p1.Y - oxis_p2.Y) < 8 Then
   Exit Sub
End If
For i% = 1 To line_for_change.lin(0).in_point(0)
'  poi(line_for_change.lin(0).in_point(i%)).data(1).data(0).data0.coordinate.X = _
        2 * line_for_change.lin(0).center(0).X - poi(line_for_change.lin(0).in_point(i%)).data(0).data0.coordinate.X
'  poi(line_for_change.lin(0).in_point(i%)).data(1).data(0).data0.coordinate.Y = _
               poi(line_for_change.lin(0).in_point(i%)).data(0).data0.coordinate.Y
Next i%
For i% = 1 To line_for_change.lin(0).in_point(0)
 line_for_change.lin(0).in_poi_coord(i%) = oxis_symmetry_point( _
        line_for_change.lin(0).in_poi_coord(i%), oxis_p1, oxis_p2)
Next i%
line_for_change.lin(0).coord(0) = line_for_change.lin(0).in_poi_coord(1)
line_for_change.lin(0).coord(1) = line_for_change.lin(0).in_poi_coord(line_for_change.lin(0).in_point(0))
 line_for_change.lin(0).center(1) = oxis_symmetry_point( _
        line_for_change.lin(0).center(1), oxis_p1, oxis_p2)
 line_for_change.move.X = line_for_change.lin(0).center(1).X - line_for_change.lin(0).center(0).X
 line_for_change.move.Y = line_for_change.lin(0).center(1).Y - line_for_change.lin(0).center(0).Y
line_for_change.direction = -line_for_change.direction
If line_for_change.direction = 1 Then
   line_for_change.rote_angle = angle_value_from_two_line(line_for_change.lin(0).in_poi_coord(1), _
       line_for_change.lin(0).in_poi_coord(2), m_poi(line_for_change.lin(0).in_point(1)).data(0).data0.coordinate, _
          m_poi(line_for_change.lin(0).in_point(2)).data(0).data0.coordinate)
Else
'   line_for_change.rote_angle = angle_value_from_two_line(line_for_change.lin(0).in_poi_coord(1), _
       line_for_change.lin(0).in_poi_coord(2), m_poi(line_for_change.lin(0).in_point(1)).data(1).data(0).data0.coordinate, _
          m_poi(line_for_change.lin(0).in_point(2)).data(1).data(0).data0.coordinate)
End If
is_first_move = False
End Sub

Public Sub oxis_symmetry_circle(oxis_p1 As POINTAPI, _
                    oxis_p2 As POINTAPI)
Dim i%
If Abs(oxis_p1.X - oxis_p2.X) + Abs(oxis_p1.Y - oxis_p2.Y) < 8 Then
   Exit Sub
End If
If m_Circ(Circle_for_change.c).data(0).data0.center > 0 Then
Circle_for_change.move = oxis_symmetry_point(m_poi( _
         m_Circ(Circle_for_change.c).data(0).data0.center).data(0).data0.coordinate, oxis_p1, _
               oxis_p2)
Circle_for_change.c_coord = oxis_symmetry_point(m_Circ(Circle_for_change.c).data(0).data0.c_coord, oxis_p1, _
               oxis_p2)
Else
Circle_for_change.move = oxis_symmetry_point(m_Circ(Circle_for_change.c).data(0).data0.c_coord, oxis_p1, _
               oxis_p2)
Circle_for_change.c_coord = oxis_symmetry_point(m_Circ(Circle_for_change.c).data(0).data0.c_coord, oxis_p1, _
               oxis_p2)
End If
For i% = 1 To m_Circ(Circle_for_change.c).data(0).data0.in_point(0)
Circle_for_change.poi_coordinate(i%) = oxis_symmetry_point(m_poi( _
         m_Circ(Circle_for_change.c).data(0).data0.in_point(i%)).data(0).data0.coordinate, oxis_p1, _
               oxis_p2)
Next i%
Circle_for_change.radii = m_Circ(Circle_for_change.c).data(0).data0.radii
Circle_for_change.direction = -1
End Sub
Public Sub center_symmetry_circle(cp As POINTAPI)
Dim i%
Circle_for_change.c_coord = center_symmetry_point(m_Circ(Circle_for_change.c).data(0).data0.c_coord, cp)
Circle_for_change.move.X = Circle_for_change.c_coord.X - _
       m_Circ(Circle_for_change.c).data(0).data0.c_coord.X
Circle_for_change.move.Y = Circle_for_change.c_coord.Y - _
       m_Circ(Circle_for_change.c).data(0).data0.c_coord.Y
For i% = 1 To m_Circ(Circle_for_change.c).data(0).data0.in_point(0)
Circle_for_change.poi_coordinate(i%) = center_symmetry_point(m_poi( _
         m_Circ(Circle_for_change.c).data(0).data0.in_point(i%)).data(0).data0.coordinate, _
               cp)
Next i%
Circle_for_change.radii = m_Circ(Circle_for_change.c).data(0).data0.radii
Circle_for_change.direction = -1
End Sub

Public Function center_symmetry_point(p As POINTAPI, cp As POINTAPI) As POINTAPI
  center_symmetry_point.X = 2 * cp.X - p.X
  center_symmetry_point.Y = 2 * cp.Y - p.Y
End Function
Public Sub center_symmetry_polygon(p As polygon, cp As POINTAPI)
Dim i%
For i% = 0 To p.total_v - 1
p.coord(i%) = center_symmetry_point(m_poi(p.v(i%)).data(0).data0.coordinate, _
      center_p)
Next i%
p.coord_center = center_symmetry_point(p.center, center_p)
 Polygon_for_change.move.X = Polygon_for_change.p(0).coord_center.X - Polygon_for_change.p(0).center.X
 Polygon_for_change.move.Y = Polygon_for_change.p(0).coord_center.Y - Polygon_for_change.p(0).center.Y
 If Polygon_for_change.rote_angle < PI Then
     Polygon_for_change.rote_angle = Polygon_for_change.rote_angle - PI
 Else
     Polygon_for_change.rote_angle = Polygon_for_change.rote_angle + PI
 End If
End Sub
Public Sub center_symmetry_line(cp As POINTAPI)
Dim i%
For i% = 1 To line_for_change.lin(0).in_point(0)
 line_for_change.lin(0).in_poi_coord(i%) = center_symmetry_point( _
        m_poi(line_for_change.lin(0).in_point(i%)).data(0).data0.coordinate, cp)
Next i%
 line_for_change.lin(0).center(1) = center_symmetry_point( _
        line_for_change.lin(0).center(0), cp)
 line_for_change.move.X = line_for_change.lin(0).center(1).X - line_for_change.lin(0).center(0).X
 line_for_change.move.Y = line_for_change.lin(0).center(1).Y - line_for_change.lin(0).center(0).Y
 If line_for_change.rote_angle < PI Then
     line_for_change.rote_angle = line_for_change.rote_angle - PI
 Else
     line_for_change.rote_angle = line_for_change.rote_angle + PI
 End If
line_for_change.lin(0).coord(0) = line_for_change.lin(0).in_poi_coord(1)
line_for_change.lin(0).coord(1) = line_for_change.lin(0).in_poi_coord(line_for_change.lin(0).in_point(0))
is_first_move = False
End Sub
Public Sub link_v_for_change_polygon()
Dim i%
For i% = 0 To Polygon_for_change.p(0).total_v - 1
Draw_form.Line (m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.X, _
 m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.Y)- _
    (Polygon_for_change.p(0).coord(i%).X, Polygon_for_change.p(0).coord(i%).Y), QBColor(7)
Next i%
End Sub
Public Sub link_v_for_change_line()
Dim i%
For i% = 1 To line_for_change.lin(0).in_point(0)
Draw_form.Line (m_poi(line_for_change.lin(0).in_point(i%)).data(0).data0.coordinate.X, _
 m_poi(line_for_change.lin(0).in_point(i%)).data(0).data0.coordinate.Y)- _
    (line_for_change.lin(0).in_poi_coord(i%).X, line_for_change.lin(0).in_poi_coord(i%).Y), QBColor(7)
Next i%
End Sub
Public Sub link_v_for_change_circle()
Dim i%
Draw_form.Line (m_Circ(Circle_for_change.c).data(0).data0.c_coord.X, _
     m_Circ(Circle_for_change.c).data(0).data0.c_coord.Y)- _
      (Circle_for_change.c_coord.X, Circle_for_change.c_coord.Y), QBColor(7)
For i% = 1 To m_Circ(Circle_for_change.c).data(0).data0.in_point(0)
Draw_form.Line (m_poi(m_Circ(Circle_for_change.c).data(0).data0.in_point(i%)).data(0).data0.coordinate.X, _
 m_poi(m_Circ(Circle_for_change.c).data(0).data0.in_point(i%)).data(0).data0.coordinate.Y)- _
    (Circle_for_change.poi_coordinate(i%).X, Circle_for_change.poi_coordinate(i%).Y), QBColor(7)
Next i%

End Sub
Public Sub draw_change_polygon(ty As Byte)
'ty=0 down ty= 1
Dim i%
Draw_form.Cls
Draw_form.DrawMode = 13
If ty = 0 Then
      Call fill_color_for_polygon(Polygon_for_change.p(0).v, Polygon_for_change.p(0).total_v, QBColor(3), 5)
      If Polygon_for_change.direction = 1 Then
      Call fill_color_for_polygon0(Polygon_for_change.p(0).coord, Polygon_for_change.p(0).total_v, QBColor(12), 1)
      Else
      Call fill_color_for_polygon0(Polygon_for_change.p(0).coord, Polygon_for_change.p(0).total_v, QBColor(10), 1)
      End If
Else
      Call fill_color_for_polygon(Polygon_for_change.p(0).v, Polygon_for_change.p(0).total_v, QBColor(3), 5)
      If Polygon_for_change.direction = 1 Then
      Call fill_color_for_polygon0(Polygon_for_change.p(0).coord, Polygon_for_change.p(0).total_v, QBColor(12), 4)
      Else
      Call fill_color_for_polygon0(Polygon_for_change.p(0).coord, Polygon_for_change.p(0).total_v, QBColor(10), 4)
      End If

End If
  For i% = 0 To Polygon_for_change.p(0).total_v - 1
   Call draw_plus_point(Draw_form, Polygon_for_change.p(0).v(i%), Polygon_for_change.p(0).coord(i%), _
            display)
  Next i%
      Call BitBlt(Draw_form.hdc, 0, 0, Draw_form.Picture1.width, _
               Draw_form.Picture1.Height, Draw_form.Picture1.hdc, 0, 0, vbSrcAnd) '把原图考贝到Draw_form
    Draw_form.DrawMode = 10
    Draw_form.fillstyle = 1
End Sub
Public Sub draw_change_line(l_c%)
Dim i%
If m_poi(line_for_change.lin(0).poi(0)).data(0).data0.coordinate.X = _
    line_for_change.lin(0).coord(0).X And _
   m_poi(line_for_change.lin(0).poi(0)).data(0).data0.coordinate.Y = _
    line_for_change.lin(0).coord(0).Y And _
   m_poi(line_for_change.lin(0).poi(1)).data(0).data0.coordinate.X = _
    line_for_change.lin(0).coord(1).X And _
   m_poi(line_for_change.lin(0).poi(1)).data(0).data0.coordinate.Y = _
    line_for_change.lin(0).coord(1).Y Then
     Exit Sub
End If
For i% = 1 To line_for_change.lin(0).in_point(0)
Call draw_plus_point(Draw_form, line_for_change.lin(0).in_point(i%), line_for_change.lin(0).in_poi_coord(i%), _
       display)
Next i%
Draw_form.Line (line_for_change.lin(0).coord(0).X, line_for_change.lin(0).coord(0).Y)- _
 (line_for_change.lin(0).coord(1).X, line_for_change.lin(0).coord(1).Y), QBColor(l_c%)
End Sub
Public Sub draw_change_circle(ty As Byte)
'ty=0 down ty= 1
'change_ty=0 ,ty= 1 中心对称 ty=2轴对称
Dim i%
Draw_form.Cls
Draw_form.DrawMode = 13
If ty = 0 Then
      Call fill_color_for_circle(m_Circ(Circle_for_change.c).data(0).data0.c_coord, _
              m_Circ(Circle_for_change.c).data(0).data0.radii, QBColor(3), 5)
     If Circle_for_change.direction = 1 Then
      Call fill_color_for_circle(Circle_for_change.c_coord, _
               Circle_for_change.radii, QBColor(12), 1)
     Else
      Call fill_color_for_circle(Circle_for_change.c_coord, _
               Circle_for_change.radii, QBColor(10), 1)
     End If
Else
      Call fill_color_for_circle(m_Circ(Circle_for_change.c).data(0).data0.c_coord, _
            m_Circ(Circle_for_change.c).data(0).data0.radii, QBColor(3), 5)
     If Circle_for_change.direction = 1 Then
      Call fill_color_for_circle(Circle_for_change.c_coord, _
               Circle_for_change.radii, QBColor(12), 4)
     Else
      Call fill_color_for_circle(Circle_for_change.c_coord, _
               Circle_for_change.radii, QBColor(10), 4)
     End If
End If
  For i% = 1 To m_Circ(Circle_for_change.c).data(0).data0.in_point(0)
   Call draw_plus_point(Draw_form, m_Circ(Circle_for_change.c).data(0).data0.in_point(i%), _
             Circle_for_change.poi_coordinate(i%), display)
  Next i%
     Call draw_plus_point(Draw_form, m_Circ(Circle_for_change.c).data(0).data0.center, _
             Circle_for_change.c_coord, display)
         Call BitBlt(Draw_form.hdc, 0, 0, Draw_form.Picture1.width, _
               Draw_form.Picture1.Height, Draw_form.Picture1.hdc, 0, 0, vbSrcAnd)
    Draw_form.DrawMode = 10
    Draw_form.fillstyle = 1
End Sub
Public Sub draw_change_circle0(fillstyle As Byte, co As Long, change_or_inform As Byte)
Dim i%
Call fill_color_for_circle(Circle_for_change.c_coord, Circle_for_change.radii, _
          co, fillstyle)
  For i% = 1 To m_Circ(Circle_for_change.c).data(0).data0.in_point(0)
   Call draw_plus_point(Draw_form, m_Circ(Circle_for_change.c).data(0).data0.in_point(i%), _
           Circle_for_change.poi_coordinate(i%), display)
  Next i%
   If m_Circ(Circle_for_change.c).data(0).data0.center > 0 Then
    Call draw_plus_point(Draw_form, m_Circ(Circle_for_change.c).data(0).data0.center, _
           Circle_for_change.c_coord, display)
   End If
End Sub
Public Sub draw_change_circle_()
Dim i%
 If Circle_for_change.direction = 1 Then
  Draw_form.Circle (Circle_for_change.c_coord.X, Circle_for_change.c_coord.Y), Circle_for_change _
          .radii, QBColor(12)
 Else
  Draw_form.Circle (Circle_for_change.c_coord.X, Circle_for_change.c_coord.Y), Circle_for_change _
          .radii, QBColor(10)
 End If
  For i% = 1 To m_Circ(Circle_for_change.c).data(0).data0.in_point(0)
   Call draw_plus_point(Draw_form, m_Circ(Circle_for_change.c).data(0).data0.in_point(i%), _
             Circle_for_change.poi_coordinate(i%), display)
  Next i%
     Call draw_plus_point(Draw_form, m_Circ(Circle_for_change.c).data(0).data0.center, _
             Circle_for_change.c_coord, display)
End Sub
 
Public Function oxis_symmetry_point(p As POINTAPI, oxis_p1 As POINTAPI, _
                    oxis_p2 As POINTAPI) As POINTAPI
Dim c1&, c2&, A1&, A2&, b1&, b2&, d&
A1& = oxis_p1.X - oxis_p2.X
b1& = oxis_p1.Y - oxis_p2.Y
A2& = b1&
b2& = -A1&
c1& = A1& * p.X + b1& * p.Y
c2& = b1& * (2 * oxis_p1.X - p.X) - A1& * (2 * oxis_p1.Y - p.Y)
d& = A1& * A1& + b1& * b1&
If d& > 0 Then
 oxis_symmetry_point.X = (b1& * c2& - b2& * c1&) / d&
 oxis_symmetry_point.Y = (-A1& * c2& + A2& * c1&) / d&
End If
End Function

Public Sub FloodFill_change_polygon()
Dim i%
Dim X&, Y&
X& = 0
Y& = 0
For i% = 0 To Polygon_for_change.p(0).total_v - 1
X& = X& + Polygon_for_change.p(0).coord(i%).X
Y& = Y& + Polygon_for_change.p(0).coord(i%).Y
Next i%
X& = X& / Polygon_for_change.p(0).total_v
Y& = Y& / Polygon_for_change.p(0).total_v
Draw_form.DrawMode = 13
Draw_form.fillstyle = 4
Call FloodFill(Draw_form.hdc, X&, Y&, QBColor(12))
Draw_form.DrawMode = 10
Draw_form.fillstyle = 1
End Sub

Public Function choce_polygon_for_change(ByVal p%) As Boolean
Dim c_x&, c_y& 'true is complete
Dim i%
If p% = Polygon_for_change.p(0).v(0) And _
     Polygon_for_change.p(0).total_v > 0 Then
Call line_number(Polygon_for_change.p(0).v(0), _
                 Polygon_for_change.p(0).v(Polygon_for_change.p(0).total_v - 1), _
                 pointapi0, pointapi0, _
                 depend_condition(0, 0), _
                 depend_condition(0, 0), _
                 conclusion, conclusion_color, 1, 0)
'Call draw_red_line(line_number5(Polygon_for_change.p.v(0), _
     Polygon_for_change.p.v(Polygon_for_change.p.total_v - 1), 0, 0, 0))
        choce_polygon_for_change = True
         set_change_fig = polygon_
          is_first_move = True
     last_conditions.last_cond(1).change_picture_type = polygon_
     MDIForm1.set_change_type.Enabled = True
     MDIForm1.set_picture_for_change.Enabled = False
     MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4090, "")
       c_x& = 0
       c_y& = 0
       For i% = 0 To Polygon_for_change.p(0).total_v - 1
       c_x& = c_x& + Polygon_for_change.p(0).coord(i%).X
       c_y& = c_y& + Polygon_for_change.p(0).coord(i%).Y
      Next i%
      c_x& = c_x& / Polygon_for_change.p(0).total_v
      c_y& = c_y& / Polygon_for_change.p(0).total_v
      Polygon_for_change.p(0).center.X = c_x&
      Polygon_for_change.p(0).center.Y = c_y&
      Polygon_for_change.p(0).coord_center.X = c_x&
      Polygon_for_change.p(0).coord_center.Y = c_y&
   For i% = 0 To Polygon_for_change.p(0).total_v - 1
'        m_poi(Polygon_for_change.p(0).v(i%)).data(1).data(0).data0.coordinate.Y = _
'         m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.Y
'    m_poi(Polygon_for_change.p(0).v(i%)).data(1).data(0).data0.coordinate.X = _
              2 * Polygon_for_change.p(0).center.X - m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.X
   Next i%
Else
'Call C_display_picture.draw_red_point(p%) 'BPset(poi(p%).data(0).data0.coordinate.X, poi(p%).data(0).data0.coordinate.Y, "", 12, display)
Polygon_for_change.p(0).coord(Polygon_for_change.p(0).total_v) = _
  m_poi(p%).data(0).data0.coordinate
  Polygon_for_change.p(0).v(Polygon_for_change.p(0).total_v) = p%
 If Polygon_for_change.p(0).total_v > 0 Then
 Call line_number(Polygon_for_change.p(0).v(Polygon_for_change.p(0).total_v - 1), _
                  Polygon_for_change.p(0).v(Polygon_for_change.p(0).total_v), _
                  pointapi0, pointapi0, _
                  depend_condition(0, 0), _
                  depend_condition(0, 0), _
                  conclusion, conclusion_color, 1, 0)
 'Call draw_red_line(line_number5( _
    Polygon_for_change.p.v(Polygon_for_change.p.total_v - 1), _
      Polygon_for_change.p.v(Polygon_for_change.p.total_v), 0, 0, 0))
End If
  Polygon_for_change.p(0).total_v = Polygon_for_change.p(0).total_v + 1
choce_polygon_for_change = False
End If
Polygon_for_change.direction = 1
End Function

Public Function choce_circle_for_change(ByVal p%, c%) As Boolean
Dim i%, j%
For i% = 1 To last_conditions.last_cond(1).circle_no
 If m_Circ(i%).data(0).data0.visible > 0 And m_Circ(i%).data(0).data0.center = p% Then
     c% = i%
      m_Circ(i%).data(0).data0.c_coord = _
        m_poi(m_Circ(i%).data(0).data0.center).data(0).data0.coordinate
       Call set_circle_color(c%, conclusion_color)
        Circle_for_change.c = c%
         Circle_for_change.move.X = 0
          Circle_for_change.move.Y = 0
          Circle_for_change.radii = m_Circ(i%).data(0).data0.radii
          Circle_for_change.c_coord = m_Circ(i%).data(0).data0.c_coord
    For j% = 1 To m_Circ(c%).data(0).data0.in_point(0)
      Circle_for_change.poi_coordinate(i%) = _
          m_poi(m_Circ(c%).data(0).data0.in_point(i%)).data(0).data0.coordinate
'      m_poi(m_circ(c%).data(0).data0.in_point(i%)).data(1).data(0).data0.coordinate.X = _
           2 * m_circ(c%).data(0).data0.c_coord.X - m_poi(m_circ(c%).data(0).data0.in_point(i%)).data(1).data(0).data0.coordinate.X
'      m_poi(m_circ(c%).data(0).data0.in_point(i%)).data(1).data(0).data0.coordinate.Y = _
           m_poi(m_circ(c%).data(0).data0.in_point(i%)).data(1).data(0).data0.coordinate.Y
    Next j%
     Circle_for_change.direction = 1
     MDIForm1.set_change_type.Enabled = True
     last_conditions.last_cond(1).change_picture_type = circle_
     MDIForm1.set_picture_for_change.Enabled = False
         choce_circle_for_change = True
          is_first_move = True
           Exit Function
 End If
Next i%
End Function
Public Function choce_line_for_change(ByVal p1%, ByVal p2%) As Boolean
Dim i%, l%
Dim n1%, n2%
l% = line_number0(p1%, p2%, n1%, n2%)
If n1% > n2% Then
 Call exchang(n1%, n2%)
End If
line_for_change.lin(0).poi(0) = m_lin(l%).data(0).in_point(n1%)
line_for_change.lin(0).poi(1) = m_lin(l%).data(0).in_point(n2%)
For i% = 1 To n2% - n1% + 1
 line_for_change.lin(0).in_point(i%) = m_lin(l%).data(0).in_point(n1% + i% - 1)
Next i%
line_for_change.lin(0).center(0).X = (m_poi(line_for_change.lin(0).poi(0)).data(0).data0.coordinate.X + _
    m_poi(line_for_change.lin(0).poi(1)).data(0).data0.coordinate.X) / 2
line_for_change.lin(0).center(0).Y = (m_poi(line_for_change.lin(0).poi(0)).data(0).data0.coordinate.Y + _
    m_poi(line_for_change.lin(0).poi(1)).data(0).data0.coordinate.Y) / 2
line_for_change.lin(0).center(1) = line_for_change.lin(0).center(0)
line_for_change.lin(0).coord(0) = m_poi(line_for_change.lin(0).poi(0)).data(0).data0.coordinate
line_for_change.lin(0).coord(1) = m_poi(line_for_change.lin(0).poi(1)).data(0).data0.coordinate
'poi(line_for_change.lin(0).poi(0)).data(1).data(0).data0.coordinate.X = _
            2 * line_for_change.lin(0).center(0).X - _
               poi(line_for_change.lin(0).poi(0)).data(0).data0.coordinate.X
'poi(line_for_change.lin(0).poi(1)).data(1).data(0).data0.coordinate.Y = _
             poi(line_for_change.lin(0).poi(1)).data(0).data0.coordinate.Y '反向
line_for_change.similar_ratio = 1
line_for_change.rote_angle = 0
line_for_change.direction = 1
 line_for_change.lin(0).in_point(0) = n2% - n1% + 1
For i% = 1 To n2% - n1% + 1
 line_for_change.lin(0).in_poi_coord(i%) = m_poi(line_for_change.lin(0).in_point(i%)).data(0).data0.coordinate
'poi(line_for_change.lin(0).in_point(i%)).data(1).data(0).data0.coordinate.X = _
            2 * line_for_change.lin(0).center(0).X - _
               m_poi(line_for_change.lin(0).in_point(i%)).data(0).data0.coordinate.X
'poi(line_for_change.lin(0).in_point(i%)).data(1).data(0).data0.coordinate.Y = _
             m_poi(line_for_change.lin(0).in_point(i%)).data(0).data0.coordinate.Y '反向
Next i%
 Call line_number(line_for_change.lin(0).poi(0), line_for_change.lin(0).poi(1), _
                  pointapi0, pointapi0, _
                  depend_condition(0, 0), _
                  depend_condition(0, 0), _
                  conclusion, conclusion_color, 1, 0)
      MDIForm1.set_picture_for_change.Enabled = False
      MDIForm1.set_change_type.Enabled = True
     last_conditions.last_cond(1).change_picture_type = line_
         choce_line_for_change = True
          set_change_fig = line_
          is_first_move = True
 End Function
Public Sub fill_color_for_polygon(n() As Integer, no As Byte, co As Long, _
             fillstyle As Byte)
              'change = 0,inform=1
Dim i%
Dim tp(6) As POINTAPI
For i% = 0 To no - 1
tp(i%) = m_poi(n(i%)).data(0).data0.coordinate
Next i%
Call fill_color_for_polygon0(tp, no, co, fillstyle)
End Sub
Public Sub fill_color_for_polygon0(coord() As POINTAPI, no As Byte, co As Long, _
             fillstyle As Byte)
Dim i%
Dim p As POINTAPI
If line_width < 2 Then
Draw_form.DrawWidth = 2
End If
Draw_form.FillColor = co
Draw_form.fillstyle = fillstyle
p.X = coord(0).X
p.Y = coord(0).Y
For i% = 0 To no - 2
 Call Drawline(Draw_form, co, 0, _
      coord(i%), coord(i% + 1), 0)
      p.X = p.X + coord(i% + 1).X
      p.Y = p.Y + coord(i% + 1).Y
Next i%
p.X = p.X / no
p.Y = p.Y / no
  Call Drawline(Draw_form, co, 0, _
      coord(0), coord(no - 1), 0)
Call FloodFill(Draw_form.hdc, p.X, _
       p.Y, co)
Draw_form.DrawWidth = line_width
End Sub
Public Sub fill_color_for_circle(center As POINTAPI, radii&, co As Long, _
             fillstyle As Byte)
Dim i%
Dim p As POINTAPI
If line_width < 2 Then
Draw_form.DrawWidth = 2
End If
Draw_form.FillColor = co
Draw_form.fillstyle = fillstyle
Draw_form.Circle (center.X, center.Y), radii, co
Draw_form.DrawWidth = line_width
Draw_form.fillstyle = 1
End Sub

Public Function angle_value_from_two_line(p1 As POINTAPI, p2 As POINTAPI, _
                     p3 As POINTAPI, p4 As POINTAPI) As Single
Dim r1!, r2!, r3!, r4!, c!
Dim Area%
Dim tp(2) As POINTAPI
tp(2).X = p4.X - p3.X + p1.X
tp(2).Y = p4.Y - p3.Y + p1.Y
tp(0) = p1
tp(1) = p2
r3! = (tp(2).X - tp(1).X) ^ 2 + (tp(2).Y - tp(1).Y) ^ 2
r1! = (tp(0).X - tp(1).X) ^ 2 + (tp(0).Y - tp(1).Y) ^ 2
r2! = (tp(0).X - tp(2).X) ^ 2 + (tp(0).Y - tp(2).Y) ^ 2
r4! = sqr(r1!) * sqr(r2!) * 2
c! = (r1! + r2! - r3!) / r4!
   If c! <> "0" Then
    c! = sqr(1 - c! ^ 2) / c!
    angle_value_from_two_line = Atn(c!)
   End If
Area% = area_triangle_from_three_point(tp(1), tp(0), tp(2))
If Area% < 0 Then
   angle_value_from_two_line = -angle_value_from_two_line
End If
End Function

Public Sub set_old_picture()
Dim i%
'For i% = 1 To last_conditions.last_cond(1).line_no
Call move_line_inner_data(0, 1)
'Next i%
'For i% = 1 To last_conditions.last_cond(1).circle_no
Call move_circle_inner_data(0, 1)
 'm_circ(i%).data(1) = Circ(i%).data(0)
'Next i%
 last_conditions.last_cond(1).con_line_no = 0
End Sub
Public Sub get_old_picture()
Dim i%
Call move_line_inner_data(1, 0)
'For i% = 1 To last_conditions.last_cond(1).circle_no
Call move_circle_inner_data(1, 0)
' Circ(i%).data(0) = Circ(i%).data(1)
'Next i%
 last_conditions.last_cond(1).con_line_no = 0
End Sub


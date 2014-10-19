Attribute VB_Name = "print_for_printer"
Option Explicit
Global start_page_no%
Global point_name_position(26) As Byte
Global page_note_string As String
Global string_at_top_of_page As String
Global p_x%, p_y%, text_y%, change_line_no%, page_no%
Global p_theorem(1 To 20) As Integer
Global last_p_theorem As Integer


Public Function area_triangle_from_three_point(p1 As POINTAPI, _
  p2 As POINTAPI, p3 As POINTAPI) As Integer
Dim Area As Long
Area = p1.X * p2.Y + p2.X * p3.Y + p3.X * p1.Y _
        - p1.X * p3.Y - p2.X * p1.Y - p3.X * p2.Y
area_triangle_from_three_point = Sgn(Area)
End Function


Public Function point_in_circle(p As POINTAPI, ce As POINTAPI, _
      r&) As Integer
Dim r_&
r_& = sqr((p.X - ce.X) ^ 2 + (p.Y - ce.Y) ^ 2)
r_& = r_& - r&
point_in_circle = Sgn(r_&)
End Function

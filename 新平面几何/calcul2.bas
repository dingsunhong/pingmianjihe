Attribute VB_Name = "calcul2"

Public Function add_v_string(v_s1 As v_string, v_s2 As v_string) As v_string
add_v_string.coord(0) = add_string(v_s1.coord(0), v_s2.coord(0), True, False)
add_v_string.coord(1) = add_string(v_s1.coord(1), v_s2.coord(1), True, False)
End Function
Public Function minus_v_string(v_s1 As v_string, v_s2 As v_string) As v_string
minus_v_string.coord(0) = minus_string(v_s1.coord(0), v_s2.coord(0), True, False)
minus_v_string.coord(1) = minus_string(v_s1.coord(1), v_s2.coord(1), True, False)
End Function
Public Function time_v_string_by_number(v_s1 As v_string, v As String) As v_string
time_v_string_by_number.coord(0) = time_string(v_s1.coord(0), v, True, False)
time_v_string_by_number.coord(1) = time_string(v_s1.coord(1), v, True, False)
End Function
Public Function divide_v_string_by_number(v_s1 As v_string, v As String) As v_string
divide_v_string_by_number.coord(0) = divide_string(v_s1.coord(0), v, True, False)
divide_v_string_by_number.coord(1) = divide_string(v_s1.coord(1), v, True, False)
End Function
'Public Function cross_time_v_string(v_s1 As v_string, v_s2 As v_string) As String
'cross_time_v_string = minus_string(time_string(v_s1.coord(0), v_s2.coord(1), False, False), _
            time_string(v_s1.coord(1), v_s2.coord(0), False, False), True, False)
'End Function
Public Function inner_time_v_string(v_s1 As v_string, v_s2 As v_string) As String
Dim t_s(2) As String
t_s(0) = time_string(v_s1.coord(0), v_s2.coord(0), True, False)
t_s(1) = add_string(time_string(v_s1.coord(0), v_s2.coord(1), False, False), _
            time_string(v_s1.coord(1), v_s2.coord(0), False, False), True, False)
t_s(2) = time_string(v_s1.coord(1), v_s2.coord(1), True, False)
inner_time_v_string = "0"
inner_time_v_string = add_string(inner_time_v_string, _
       time_string(t_s(0), "R", True, False), True, False)
inner_time_v_string = add_string(inner_time_v_string, _
       time_string(t_s(1), "S", True, False), True, False)
inner_time_v_string = add_string(inner_time_v_string, _
       time_string(t_s(2), "T", True, False), True, False)
End Function
Public Function from_v_string_to_string(v_s As v_string) As String
   from_v_string_to_string = add_string(time_string(v_s.coord(0), "U", False, False), _
       time_string(v_s.coord(1), "V", False, False), True, False)
End Function
Public Function from_string_to_v_string(ByVal v As String) As v_string
from_string_to_v_string.coord(0) = read_para_from_string_for_ietm(v, "U", v)
from_string_to_v_string.coord(1) = read_para_from_string_for_ietm(v, "V", v)
End Function
Public Function cross_time_v_string(ByVal s1$, ByVal s2$) As String
Dim v1(4) As String
Dim v2(4) As String
Dim ty(1) As Byte
v1(0) = s1
v2(0) = s2
ty(0) = string_type(v1(0), v1(1), v1(2), v1(3), v1(4))
ty(1) = string_type(v2(0), v2(1), v2(2), v2(3), v2(4))
If ty(0) = 3 Then
   cross_time_v_string = cross_time_v_string(v1(2), v2(0))
    cross_time_v_string = divide_string(cross_time_v_string, v1(3), True, False)
ElseIf ty(1) = 3 Then
    cross_time_v_string = time_string("-1", cross_time_v_string(v2(0), v1(0)), False, False)
Else
   If v1(4) <> "" Then
      cross_time_v_string = add_string(cross_time_v_string(v1(1), v2(0)), _
             cross_time_v_string(v1(4), v2(0)), True, False)
   ElseIf v2(4) <> "" Then
    cross_time_v_string = time_string("-1", cross_time_v_string(v2(0), v1(0)), False, False)
   Else
    cross_time_v_string = time_string(v1(2), v2(2), False, False)
     cross_time_v_string = time_string(cross_time_v_string, _
               cross_time_v_item(v1(3), v2(3)), True, False)
   End If
End If
End Function


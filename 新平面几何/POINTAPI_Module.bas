Attribute VB_Name = "POINTAPI_Module"
Option Explicit
Type POINTAPI
X As Long
Y As Long
End Type
Type POINTAPI_VALUE
X As String
Y As String
End Type
Public Function inner_time_two_POINTAPI(p1 As POINTAPI, p2 As POINTAPI) As Long
 inner_time_two_POINTAPI = p1.X * p2.X + p1.Y * p2.Y
End Function
Public Function add_POINTAPI(pA1 As POINTAPI, pA2 As POINTAPI) As POINTAPI
add_POINTAPI.X = pA1.X + pA2.X
add_POINTAPI.Y = pA1.Y + pA2.Y
End Function
Public Function minus_POINTAPI(pA1 As POINTAPI, pA2 As POINTAPI) As POINTAPI
minus_POINTAPI.X = pA1.X - pA2.X
minus_POINTAPI.Y = pA1.Y - pA2.Y
End Function
Public Function time_POINTAPI(pA1 As POINTAPI, pA2 As POINTAPI) As Long
time_POINTAPI = pA1.X * pA2.X + pA1.Y * pA2.Y
End Function
Public Function abs_POINTAPI(pA As POINTAPI) As Integer
abs_POINTAPI = sqr(pA.X * pA.X + pA.Y * pA.Y)
End Function
Public Function verti_POINTAPI(pA As POINTAPI) As POINTAPI
verti_POINTAPI.X = pA.Y
verti_POINTAPI.Y = -pA.X
End Function
Public Function time_POINTAPI_by_number(pA As POINTAPI, nu!) As POINTAPI
time_POINTAPI_by_number.X = pA.X * nu!
time_POINTAPI_by_number.Y = pA.Y * nu!
End Function
Public Function divide_POINTAPI_by_number(pA As POINTAPI, nu!) As POINTAPI
divide_POINTAPI_by_number.X = pA.X / nu!
divide_POINTAPI_by_number.Y = pA.Y / nu!
End Function
Public Function cross_time_POINTAPI(pA1 As POINTAPI, pA2 As POINTAPI) As Long
cross_time_POINTAPI = pA1.X * pA2.Y - pA1.Y * pA2.X
End Function
Function distance_point_to_line0(p As POINTAPI, p1 As POINTAPI, p2 As POINTAPI, dis&, _
                                   vertical_foot As POINTAPI, _
                                       Optional ty As Integer = 0) As Boolean
                                   'ty=0,ȫ��=1ֻ�������,=2��ֻ���㴹��
Dim l&, s&
Dim ratio!
Dim p12_coord As POINTAPI
Dim p02_coord As POINTAPI
l& = distance_of_two_POINTAPI(p1, p2)
If l& = 0 Then
  distance_point_to_line0 = False
Else
   p12_coord = minus_POINTAPI(p1, p2)
   p02_coord = minus_POINTAPI(p, p2)
   If ty = 0 Or ty = 1 Then
     s& = cross_time_POINTAPI(p12_coord, p02_coord)     '�൱�ڼ���ƽ���ı��ε����
      dis& = s& / l& '��������ʾ����
   End If
   If ty = 0 Or ty = 2 Then
     ratio! = inner_time_two_POINTAPI(p12_coord, p02_coord) / l& ^ 2
    vertical_foot = add_POINTAPI(p2, time_POINTAPI_by_number(p12_coord, ratio))
   End If
  distance_point_to_line0 = True
End If
End Function
Function distance_point_to_line(p As POINTAPI, Start_po As POINTAPI, parall_or_vertical As Integer, _
    p1 As POINTAPI, p2 As POINTAPI, d&, vertical_foot As POINTAPI, Optional ty As Integer = 0) As Boolean
     'p�㵽��start_po(ƽ�л�ֱ)ֱ�ߵľ���
If parall_or_vertical = paral_ Then
   distance_point_to_line = distance_point_to_line0(p, _
              add_POINTAPI(minus_POINTAPI(p2, p1), Start_po), Start_po, d&, vertical_foot)
Else
    distance_point_to_line = distance_point_to_line0(p, _
              add_POINTAPI(verti_POINTAPI(minus_POINTAPI(p2, p1)), Start_po), Start_po, d&, vertical_foot)
End If
End Function
Public Function is_point_on_line(point0 As POINTAPI, point1 As POINTAPI, point2 As POINTAPI, _
                    out_point As POINTAPI, aid_line_end_point_coord1 As POINTAPI, _
                       aid_line_end_point_coord2 As POINTAPI, Optional line_type As Integer = 0) As Integer '0 ������,1,����2���߶���,3�����߶���
Dim dis&
Dim in_coord As POINTAPI
in_coord = point0
  If distance_point_to_line0(in_coord, point1, point2, dis&, out_point, 0) Then '����㵽ֱ�ߵľ��룬���жϵ��Ƿ���ֱ����
    If Abs(dis&) < 5 Then '��ֱ����
      If line_type = aid_condition Or line_type = paral_ Or line_type = verti_ Then 'ֱ�ߵ�����
         aid_line_end_point_coord2 = out_point '�յ�
         aid_line_end_point_coord1 = point1 '���
         is_point_on_line = point_out_segement
      Else
         aid_line_end_point_coord2 = out_point
       If line_type <= condition Then
       If point1.X < point2.X Then
          If out_point.X < point1.X Then
               aid_line_end_point_coord1 = point1
             is_point_on_line = point_out_segement
          ElseIf point2.X < out_point.X Then
             aid_line_end_point_coord1 = point2
             is_point_on_line = point_out_segement
          Else
             is_point_on_line = point_on_segement
          End If
       ElseIf point1.X > point2.X Then
          If out_point.X > point1.X Then
             aid_line_end_point_coord1 = point1
             is_point_on_line = point_out_segement
          ElseIf point2.X > out_point.X Then
             aid_line_end_point_coord1 = point2
             is_point_on_line = point_out_segement
          Else
             is_point_on_line = point_on_segement
          End If
       ElseIf point1.Y < point2.Y Then
          If out_point.Y < point1.Y Then
             aid_line_end_point_coord1 = point1
             is_point_on_line = point_out_segement
          ElseIf point2.Y < out_point.Y Then
             aid_line_end_point_coord1 = point2
             is_point_on_line = point_out_segement
          Else
             is_point_on_line = point_on_segement
          End If
       ElseIf point2.Y < point1.Y Then
          If out_point.Y < point2.Y Then
             aid_line_end_point_coord1 = point2
             is_point_on_line = point_out_segement
          ElseIf point1.Y < out_point.Y Then
             aid_line_end_point_coord1 = point1
             is_point_on_line = point_out_segement
          Else
             is_point_on_line = point_on_segement
          End If
       End If
  '*************************************************************************
      ElseIf line_type = 2 Or line_type = 3 Or line_type = 4 Then
                aid_line_end_point_coord2 = point1
             is_point_on_line = point_out_segement
     End If
      End If
   Else
       is_point_on_line = point_not_on_line
       aid_line_end_point_coord1.X = 10000
       aid_line_end_point_coord1.Y = 10000
       aid_line_end_point_coord2.X = 10000
       aid_line_end_point_coord2.Y = 10000
   End If
  End If
End Function
Public Function distance_of_two_POINTAPI(pA1 As POINTAPI, pA2 As POINTAPI) As Long
Dim r!
'On Error GoTo dis_of_two_point_end
Dim dis_pointapi As POINTAPI
dis_pointapi = minus_POINTAPI(pA1, pA2)
r! = dis_pointapi.X ^ 2 + dis_pointapi.Y ^ 2
distance_of_two_POINTAPI = sqr(r!)
End Function
Public Function mid_POINTAPI(pA1 As POINTAPI, pA2 As POINTAPI) As POINTAPI
mid_POINTAPI.X = (pA1.X + pA2.X) / 2
mid_POINTAPI.Y = (pA1.Y + pA2.Y) / 2
End Function
Public Function devite_two_POINTAPI_by_ratio(p1 As POINTAPI, p2 As POINTAPI, r!) As POINTAPI
 devite_two_POINTAPI_by_ratio = add_POINTAPI(time_POINTAPI_by_number(minus_POINTAPI(p2, p1), r!), p1)
End Function

Public Function is_same_POINTAPI(p_coord1 As POINTAPI, p_coord2 As POINTAPI) As Boolean
  If distance_of_two_POINTAPI(p_coord1, p_coord2) < 6 Then
     is_same_POINTAPI = True
  End If
End Function


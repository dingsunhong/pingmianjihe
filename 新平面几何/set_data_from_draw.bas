Attribute VB_Name = "set_data_from_draw"
Option Explicit
Public Function set_inter_point_line_line_data(ByVal p1%, ByVal ty1 As Byte, ByVal l1%, ByVal p2%, _
                                                     ByVal ty2 As Byte, ByVal l2%, out_l1%, out_l2%, out_p%, _
                                                           c_data As condition_data_type) As Byte
Dim paral_or_verti_string(1) As String
Dim temp_record As total_record_type
If ty1 = True Then
    paral_or_verti_string(0) = "平行于"
Else
    paral_or_verti_string(0) = "垂直于"
End If
If ty2 = True Then
    paral_or_verti_string(1) = "平行于"
Else
    paral_or_verti_string(1) = "垂直于"
End If
Call add_record_to_record(c_data, temp_record.record_data.data0.condition_data)
Call add_point_to_line(out_p%, out_l1%, 0, True, True, 0, temp_record)
Call add_point_to_line(out_p%, out_l2%, 0, True, True, 0, temp_record)
  If p1% = 0 And p2% = 0 Then
    m_poi(out_p%).data(0).inform = "直线" + m_lin(l1%).data(0).inform + "与" + _
                                   "直线" + m_lin(l2%).data(0).inform + "的交点"
  ElseIf p2% = 0 And p1% > 0 Then
   m_poi(out_p%).data(0).inform = "过" + paral_or_verti_string(0) + m_lin(l1%).data(0).inform + "的直线与" + _
                                  "直线" + m_lin(l2%).data(0).inform + "的交点"
   If ty1 = True Then
    set_inter_point_line_line_data = set_dparal(l1%, out_l1%, temp_record, 0, 0, 0)
   Else
    set_inter_point_line_line_data = set_dverti(l1%, out_l1%, temp_record, 0, 0, 0)
   End If
  ElseIf p2% > 0 And p1% > 0 Then
   m_poi(out_p%).data(0).inform = "过" + m_poi(p1%).data(0).data0.name + _
                                  paral_or_verti_string(0) + m_lin(l1%).data(0).inform + "的直线与" + _
                                 "过" + m_poi(p2%).data(0).data0.name + _
                                  paral_or_verti_string(1) + m_lin(l2%).data(0).inform + "的直线与"
   If ty1 = True Then
    set_inter_point_line_line_data = set_dparal(l1%, out_l1%, temp_record, 0, 0, 0)
   Else
    set_inter_point_line_line_data = set_dverti(l1%, out_l1%, temp_record, 0, 0, 0)
   End If
   If ty2 = True Then
    set_inter_point_line_line_data = set_dparal(l2%, out_l2%, temp_record, 0, 0, 0)
   Else
    set_inter_point_line_line_data = set_dverti(l2%, out_l2%, temp_record, 0, 0, 0)
   End If
  End If
If run_type = 5 Then
    set_inter_point_line_line_data = set_New_point(out_p%, temp_record, out_l1%, out_l2%, 0, 0, 0, 0, 0, 1)
End If
End Function



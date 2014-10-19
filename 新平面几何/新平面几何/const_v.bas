Attribute VB_Name = "const_v"
Option Explicit

Type V_line_value_data0_type
 v_line As Integer
  value As String
   squre_value As String
    line_value_no As Integer
     record As record_data_type
End Type
Type V_line_value_type
data(8) As V_line_value_data0_type
record_ As record_type
End Type
Global V_line_value() As V_line_value_type
Global con_V_line_value(3) As V_line_value_type
Type V_two_line_time_value_data0_type
 poi(3) As Integer
 lin(1) As Integer
 n(3) As Integer
 value As Integer
 record As record_data_type
End Type
Type V_two_line_time_value_type
data(8) As V_two_line_time_value_data0_type
record_ As record_type
End Type
Global V_two_line_time_value() As V_line_value_type
Global con_V_two_line_time_value(3) As V_two_line_time_value_type



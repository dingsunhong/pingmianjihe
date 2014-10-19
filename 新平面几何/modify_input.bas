Attribute VB_Name = "modify_input"
Option Explicit
Global temp_problem_record As problem_record
Global modify_input_statue As Boolean
Global modify_input_statue_no As Integer

Public Sub back_input_to(w_no%)
Dim i%, j%, tn%
Dim temp_line As line_data_type
Dim temp_circ As circle_data_type
Call clear_wenti_display
Call init_conditions(0)
last_conditions.last_cond(1).point_no = 0
For i% = 1 To C_display_wenti.m_last_input_wenti_no
 wenti_cond0.data = temp_problem_record.wenti_cond(i%)
 wenti_cond0.data.wenti_no = i%
  Call C_display_wenti.Set_wenti
  For j% = 0 To 10
   If C_display_wenti.m_point_no(i%, j%) > 0 And C_display_wenti.m_point_no(i%, j%) <= 26 Then
    last_conditions.last_cond(1).point_no = C_display_wenti.m_point_no(i%, j%)
   End If
  Next j%
Next i%
For i% = 1 To last_conditions.last_cond(1).point_no
Call set_point_data_from_input(temp_problem_record.poi(i%))
'Call draw_point(Draw_form, poi(i%), 0, True)
Next i%
last_conditions.last_cond(1).line_no = 0
For i% = 1 To temp_problem_record.last_line
temp_line.data0.total_color = 0
For j% = 0 To 10
temp_line.in_point(i%) = 0
Next j%
temp_line.data0.poi(0) = 0
temp_line.data0.poi(1) = 0
temp_line.data0.type = 0
temp_line.data0.visible = 0
For j% = 1 To 10
 If temp_problem_record.line_no(i%).line_data.in_point(j%) <= last_conditions.last_cond(1).point_no Then
 temp_line.data0.in_point(0) = temp_line.data0.in_point(0) + 1
 temp_line.in_point(0) = temp_line.in_point(0) + 1
  temp_line.in_point(temp_line.in_point(0)) = temp_problem_record.line_no(i%).line_data.in_point(j%)
  temp_line.data0.in_point(temp_line.in_point(0)) = Abs(temp_problem_record.line_no(i%).line_data.in_point(j%))
 End If
Next j%
temp_line.data0.total_color = temp_problem_record.line_no(i%).line_data.total_color
temp_line.data0.poi(0) = temp_problem_record.line_no(i%).line_data.in_point(1)
temp_line.data0.poi(1) = temp_problem_record.line_no(i%).line_data.in_point(temp_problem_record.line_no(i%).line_data.in_point(0))
temp_line.data0.type = temp_problem_record.line_no(i%).line_data.type
temp_line.data0.visible = temp_problem_record.line_no(i%).line_data.visible
last_conditions.last_cond(1).line_no = last_conditions.last_cond(1).line_no + 1
 tn% = last_conditions.last_cond(1).line_no
  If temp_line.data0.visible = 0 Then
   Call set_line_data0(tn%, temp_line, 0, 0)
  Else
   Call set_line_data0(tn%, temp_line, 0, 0)
  End If
Next i%
last_conditions.last_cond(1).circle_no = 0
For i% = 1 To temp_problem_record.last_circle
temp_circ.data0.color = 0
For j% = 0 To 10
temp_circ.data0.in_point(i%) = 0
Next j%
temp_circ.data0.c_coord.X = 0
temp_circ.data0.c_coord.Y = 0
temp_circ.data0.center = 0
temp_circ.data0.name = ""
temp_circ.data0.radii = 0

For j% = 1 To 10
 If temp_problem_record.circ(i%).circle_data.in_point(j%) <= last_conditions.last_cond(1).point_no Then
 temp_circ.data0.in_point(0) = temp_circ.data0.in_point(0) + 1
  temp_circ.data0.in_point(temp_circ.data0.in_point(0)) = temp_problem_record.circ(i%).circle_data.in_point(j%)
 End If
Next j%
temp_circ.data0.color = temp_problem_record.circ(i%).circle_data.color
temp_line.data0.type = temp_line.data0.type
temp_circ.data0.c_coord = temp_problem_record.circ(i%).circle_data.c_coord
If temp_problem_record.circ(i%).circle_data.center <= last_conditions.last_cond(1).point_no Then
temp_circ.data0.center = temp_problem_record.circ(i%).circle_data.center
End If
temp_circ.data0.name = temp_problem_record.circ(i%).circle_data.name
temp_circ.data0.radii = temp_problem_record.circ(i%).circle_data.radii
temp_circ.data0.visible = temp_problem_record.circ(i%).circle_data.visible
'If last_conditions.last_cond(1).circle_no Mod 10 = 0 Then
'ReDim Preserve Circ(last_conditions.last_cond(1).circle_no + 10) As circle_type
'End If
'last_conditions.last_cond(1).circle_no = last_conditions.last_cond(1).circle_no + 1
 Call Set_m_circle_data(0, temp_circ)
Next i%
'Call C_display_wenti.input_m_sentences(Wenti_form.Picture1, 1, 1, C_display_wenti.m_last_input_wenti_no, 0, False, 0)
For i% = 1 To C_display_wenti.m_last_input_wenti_no

Call draw_picture(i%, 0, True)
Next i%
End Sub

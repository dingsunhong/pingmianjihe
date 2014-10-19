Attribute VB_Name = "initial"
Option Explicit
Global C_wait_for_aid_point As wait_for_aid_point
Global C_display_wenti As display_class
Global C_display_wenti1 As display_class '信息库使用
Global C_curve As curve_Class
Global C_display_picture As display_picture
Global C_condition_tree As condition_tree
Global C_IO As IO_class
Global draw_statue As Byte
Dim t_line(7) As Integer
Public Sub control_menu(Enab As Boolean)
Dim i%
   For i% = 2 To 20
    If i% <> 7 Then
     MDIForm1.Toolbar1.Buttons(i%).Enabled = Enab
    End If
   Next i%
   MDIForm1.edit.Enabled = Enab
   MDIForm1.draw.Enabled = Enab
   MDIForm1.new.Enabled = Enab
   MDIForm1.Open.Enabled = Enab
   MDIForm1.save.Enabled = Enab
   MDIForm1.save_as.Enabled = Enab
   MDIForm1.savee.Enabled = Enab
   MDIForm1.mprint.Enabled = Enab
   MDIForm1.Inputcond.Enabled = Enab
   MDIForm1.conclusion.Enabled = Enab
   MDIForm1.mea_and_cal.Enabled = Enab
   MDIForm1.solve.Enabled = Enab
   MDIForm1.examp.Enabled = Enab
   MDIForm1.dbase.Enabled = Enab
End Sub
Public Sub init_project_(ch As Boolean)
Dim i%
If ch = True Then
For i% = 7 To 16
MDIForm1.Toolbar1.Buttons(i%).Image = i% + 12
Next i%
MDIForm1.Toolbar1.Buttons(20).Image = 30
MDIForm1.Toolbar1.Buttons(21).Image = 33
MDIForm1.Toolbar1.Buttons(19).Image = 29
Wenti_form.Picture1.font.Weight = 700
Wenti_form.Picture2.font.Weight = 700
inform.TreeView1.font.Weight = 700
inform.List1.font.Weight = 700
inform.Picture1.font.Weight = 700
exam_form.List1.font.Weight = 700
Wenti_form.TreeView1(0).font.Weight = 700
MDIForm1.StatusBar1.font.Weight = 700
'MDIForm1.projector.Checked = True
regist_data.projector_Checked = True
line_width = 3
Draw_form.DrawWidth = 3
Else
For i% = 7 To 19
MDIForm1.Toolbar1.Buttons(i%).Image = i% - 2
Next i%
MDIForm1.Toolbar1.Buttons(20).Image = 31
MDIForm1.Toolbar1.Buttons(21).Image = 34
Wenti_form.Picture1.font.Weight = 400
Wenti_form.Picture2.font.Weight = 400
inform.TreeView1.font.Weight = 400
inform.List1.font.Weight = 400
inform.Picture1.font.Weight = 400
exam_form.List1.font.Weight = 400
Wenti_form.TreeView1(0).font.Weight = 400
MDIForm1.StatusBar1.font.Weight = 400
'MDIForm1.projector.Checked = False
regist_data.projector_Checked = False
End If

End Sub
Public Sub init_tree_(ch As Boolean)
 If ch = True Then
 inform.TreeView1.visible = True
  inform_treeview_visible = True
   inform_picture_visible = False
    inform.VScroll1.visible = False
'     MDIForm1.treeview.Checked = True
      regist_data.treeview_Checked = True
 Else
 inform.TreeView1.visible = False
  inform_treeview_visible = False
'    MDIForm1.treeview.Checked = False
     regist_data.treeview_Checked = False
      inform.VScroll1.visible = True
       inform_picture_visible = True
 End If
End Sub

Public Sub init_input0()
Dim i%, j%
For i% = 1 To C_display_wenti.m_last_input_wenti_no
 For j% = 0 To 10
  If C_display_wenti.m_condition(i%, j%) >= "A" And _
       C_display_wenti.m_condition(i%, j%) <= "Z" Then
        Call C_display_wenti.set_m_condition(i%, global_icon_char, j%)
  End If
 Next j%
Next i%
End Sub
Public Sub init_wenti(no%)
Dim j%
For j% = 0 To 50 'Call init_input_cond(i%)
Call C_display_wenti.set_m_condition(no%, empty_char, j%)
Call C_display_wenti.set_m_point_no(no, 0, j%, False)
Next j%
End Sub

Public Sub init_color()
widthform.Label1.BackColor = QBColor(15)
widthform.Label2.BackColor = QBColor(15)
widthform.Label3.BackColor = QBColor(15)
widthform.Label4.BackColor = QBColor(15)
width_set_statue = 0
End Sub
Public Sub init_condition()
Dim i%, j%
yidian_type = 0
select_wenti_no% = 0
yidian_stop = True
prove_result = 0
input_last_point = 0
operator = ""
'old_last_general_string_combine = 0
'old_last_angle3_combine = 0
Polygon_for_change.p(0).total_v = 0
Polygon_for_change.similar_ratio = 1
Polygon_for_change.rote_angle = 0
Polygon_for_change.direction = 1
Polygon_for_change.move.X = 0
Polygon_for_change.move.Y = 0
For i% = 0 To 3
conclusion_data(i%).no(0) = 0
conclusion_data(i%).no(1) = 0
conclusion_data(i%).ty = 0
conclusion_data(i%).wenti_no = 0
Next i%
last_conditions.last_cond(1).pass_word_for_teacher = "*****"
area_of_triangle_conclusion = 0
line3_value_conclusion = 0
line3_angle_conclusion = 0
'new_result_from_add = False
ge_reduce_level = 0
finish_prove = 0
re_name_ty = 0
run_type = run_type_1
Call init_reduce
start_no% = 0
contro_process = 0
last_length = 0
last_length_point_to_line = 0
last_angle_value_for_measur = 0
last_Area_polygon = 0
last_constant = 0
'c_display_wenti.m_last_conclusion% = 0
event_statue = ready
modify_condition_no = -1
condition_no = -1
input_text_statue = False
is_uselly_para_for_angle = True
is_uselly_degree_for_angle = True
ruler_display = False
prove_times = 0
last_add_condition = 0
last_conclusion = 0
last_measur_string = 0
Call remove_init
Call init_record
input_last_point = 0
display_no = 0
Erase display_string
prove_step = 0
Last_hotpoint_of_theorem1 = 0
Erase Hotpoint_of_theorem1
last_prove_by_hand_no = 0
last_con_line1 = 0
'turn_over_type = 1
chose_total_theorem = False
'last_conditions.last_cond(1).line_no_for_aid = 0
'last_conditions.last_cond(1).line_no_from_two_point(0) = 0
'last_conditions.last_cond(1).line_no_from_two_point(1) = 0
last_aid_point = 0
'last_Rtriangle = 0
'last_conditions.last_cond(1).last_long_squre_no = 0
last_addition_condition = 0
last_condition_eline = 0
'last_condition_eangle = 0
'last_condition_Dpoint_pair = 0
'Last_condition_two_angle_value = 0
'last_condition_total_equal_triangle = 0
'last_condition_similar_triangle = 0
'last_condition_angle3_value = 0
input_char_no = 0
last_char = 0
input_last_point = 0
'total_condition = 0
'last_two_line_value_with_line_value = 0
'last_two_line_value_with_eline = 0
'last_two_line_value_with_midpoint = 0
'last_two_line_value_with_relation = 0
'last_conditions.last_cond(1).line_no3_value_with_line_value = 0
'last_conditions.last_cond(1).line_no3_value_with_eline = 0
'last_conditions.last_cond(1).line_no3_value_with_midpoint = 0
'last_conditions.last_cond(1).line_no3_value_with_relation = 0
'last_conditions.last_cond(1).line_no_value_with_line2 = 0
'last_conditions.last_cond(1).line_no_value_with_line3 = 0
'last_eline_with_line2 = 0
'last_eline_with_line3 = 0
'last_mid_point_with_line2 = 0
'last_mid_point_with_line3 = 0
'last_relation_with_line2 = 0
'last_relation_with_line3 = 0
'last_conditions.last_cond(1).line_no3_value_with_three = 0
temp_th_ch_51 = 0
temp_th_ch_52 = 0 'write_wenti_no = 0
'Call init_theorem
End Sub
Public Sub init_conditions(t%)
Dim i%, j%
Call C_wait_for_aid_point.init
Call C_display_wenti.init(Wenti_form.Picture1, 0)
Call C_display_picture.init(Draw_form)
Call C_curve.Class_Init
save_statue = 0
If t% = 0 Then
path_and_file = ""
wenti_data_type = 0
'protect_munu = 0
Draw_form.Caption = LoadResString_(2005, "") + "-" + LoadResString_(1925, "")
 Wenti_form.Caption = LoadResString_(1960, "") + "-" + LoadResString_(1925, "") + _
                         LoadResString_(3955, "\\1\\" + LoadResString_(425, ""))
MDIForm1.Timer1.interval = 500
MDIForm1.Text1.text = ""
MDIForm1.Text1.visible = False
MDIForm1.Text2.text = ""
MDIForm1.Text2.visible = False
MDIForm1.re_name.Enabled = True
'MDIForm1.moldy_condition.Enabled = True
'MDIForm1.moldy_conclusion.Enabled = True
Wenti_form.Picture1.top = 360
Wenti_form.Picture1.left = 0
Wenti_form.HScroll1.top = Wenti_form.ScaleHeight - 68
Wenti_form.VScroll1.Height = Wenti_form.ScaleHeight - 68
Wenti_form.HScroll1.left = 0
Wenti_form.HScroll1.value = 0
Wenti_form.HScroll1.visible = True
display_wenti_h_position% = 0
'MDIForm1.Timer1.Enabled = True
'MDIForm1.method.Enabled = False
'MDIForm1.method1.Enabled = False
'MDIForm1.method2.Enabled = False
'MDIForm1.method3.Enabled = False
Draw_form.Timer1.Enabled = True
MDIForm1.set_picture_for_change.Enabled = True
MDIForm1.set_change_type.Enabled = False
MDIForm1.Inputcond.Enabled = True
MDIForm1.conclusion.Enabled = True
MDIForm1.examp.Enabled = True
MDIForm1.edit.Enabled = True
MDIForm1.Inputcond.Enabled = True
MDIForm1.draw.Enabled = True
MDIForm1.conclusion.Enabled = True
MDIForm1.mea_and_cal.Enabled = True
MDIForm1.c_line1_5.Enabled = c_line1_5_enabled
MDIForm1.c_cal.Enabled = True
MDIForm1.c_choose.Enabled = c_choose_enabled
input_text_statue = False
MDIForm1.Toolbar1.Buttons(19).ToolTipText = LoadResString_(1935, "")
MDIForm1.Toolbar1.Buttons(20).Image = 32
MDIForm1.Toolbar1.Buttons(21).Image = 34
If regist_data.line_width > 0 Then
condition_color = regist_data.condition_color
conclusion_color = regist_data.conclusion_color
fill_color = regist_data.fill_color
line_width = regist_data.line_width
Else
condition_color = 3
conclusion_color = 12
fill_color = 7
line_width = 1
Call set_regist
End If
Draw_form.DrawWidth = line_width
Draw_form.Picture1.DrawWidth = line_width
'Draw_form.Picture1.DrawWidth = line_width
'Draw_form.Picture1.Height = Draw_form.Height
'Draw_form.Picture1.Width = Draw_form.Width
'Draw_form.Picture2.DrawWidth = line_width
'Draw_form.Picture2.Height = Draw_form.Height
'Draw_form.Picture2.Width = Draw_form.Width
Unload Print_Form
'***********************************************************
Ratio_for_measure.Ratio_for_measure = 0
Ratio_for_measure.ratio_for_measure0 = 0
Ratio_for_measure.is_fixed_ratio = False
Ratio_for_measure.sons.last_son = 0
'***********************************************************
run_type = 0
inform.Picture1.Cls
inform.List1.Clear
set_change_type_ = False
modify_input_statue = False
modify_input_statue_no = 0
Call init_inform
pro_no1% = 0
run_statue = 0
yidian_type = 0
yidian_no = 0
error_of_wenti = 0
is_set_function_data = 0
For i% = 1 To 30
operate_step(i%).last_point = 0
operate_step(i%).last_con_circle = 0
operate_step(i%).last_con_line = 0
operate_step(i%).last_conclusion = 0
Next i%
conclusion_point(0).poi(0) = 0
conclusion_point(1).poi(0) = 0
conclusion_point(2).poi(0) = 0
conclusion_point(3).poi(0) = 0
wenti_no_% = 0
'old_wenti_no_% = 0
wenti_type = 0
'wenti_type0 = 0
draw_or_prove = 0
set_or_prove = 0
draw_wenti_no = 0
last_used_char = 0
draw_step = -1
write_wenti = False
record0.record_data.data0.condition_data.condition_no = 0
last_area_element_in_conclusion = 0
'Wenti_form.Text1.visible = False
Wenti_form.Label1.visible = False
Wenti_form.HScroll2.visible = False
Wenti_form.SSTab1.Tab = 0
set_change_fig = 0
operat_is_acting = False
MDIForm1.StatusBar1.Panels(1).text = ""
End If
Wenti_form.VScroll1.value = 0
'display_wenti_v_position% = 0
Draw_form.Cls
Wenti_form.Picture1.Cls
Wenti_form.Picture2.Cls
Wenti_form.TreeView1(0).Nodes.Clear
Draw_form.Picture1.visible = False
Draw_form.Picture1.Cls
Set C_display_picture = Nothing
Set C_display_picture = New display_picture
Call C_display_picture.init(Draw_form)
Draw_form.DrawMode = 10
'Draw_form.DrawWidth = 1
Draw_form.fillstyle = 1
is_new_result = False
picture_copy = False
start_page_no% = 0
last_node_index = 0
conclusion_no_wenti = 255
'*************************
For i% = 0 To 15
 red_line(i%) = 0
Next i%
For i% = 0 To 7
 fill_color_line(i%) = 0
addition_condition(i%) = 0
temp_last_point(i%) = 0
 temp_point(i%).no = 0
Next i%
For i% = 0 To 3
temp_line(i%) = 0
temp_circle(i%) = 0
Next i%
For i% = 1 To 26
temp_record_poi(i%, 0) = 0
temp_record_poi(i%, 1) = 0
Next i%
'************************
Erase temp_wenti_cond
'Erase wenti_cond
For i% = -6 To 180
th_chose(i%).used = 0
Next i%
old_point = 0
last_condition = 0
Call set_picture_data_init
For i% = -6 To 180
 th_chose(i%).chose = regist_data.th_chose(i%)
Next i%
th_chose(-6).chose = 0
th_chose(-5).chose = 0
last_problem_input% = 0
last_combine_length_of_polygon_with_line_value(0) = 0
last_combine_length_of_polygon_with_two_line_value(0) = 0
last_combine_length_of_polygon_with_line3_value(0) = 0
last_combine_length_of_polygon_with_line_value(1) = 0
last_combine_length_of_polygon_with_two_line_value(1) = 0
last_combine_length_of_polygon_with_line3_value(1) = 0
c_data_for_reduce.condition_no = 0
Call init_condition
End Sub
Public Sub init_record()
Dim i%
Erase branch_data
ReDim Preserve branch_data(0) As branch_data_type
last_conditions_for_aid_no = 0
Erase last_conditions_for_aid
last_add_aid_point_for_two_line = 0
Erase add_aid_point_for_two_line_
last_add_aid_point_for_two_circle = 0
Erase add_aid_point_for_two_circle_
last_add_aid_point_for_line_circle = 0
Erase add_aid_point_for_line_circle_
last_add_aid_point_for_mid_point = 0
Erase add_aid_point_for_mid_point_
last_add_aid_point_for_eline = 0
Erase add_aid_point_for_eline_
last_add_conditions = 0
Erase add_condition
Call init_condition_no(last_conditions.last_cond(0))
Call init_condition_no(last_conditions.last_cond(1))
Call init_condition_no(last_conditions.last_cond(2))
Call init_condition_no(t_condition.last_cond(1))
'**********
Erase branch_data
Erase Dvalue_string
Erase four_sides_fig
 ReDim four_sides_fig(0) As four_sides_fig_type
Erase angle3_value
 ReDim angle3_value(0) As angle3_value_type
 con_angle3_value(0).data(0).data0 = angle3_value(0).data(0).data0
 con_angle3_value(0).data(1).data0 = angle3_value(0).data(0).data0
 con_angle3_value(1) = con_angle3_value(0)
 con_angle3_value(2) = con_angle3_value(0)
 con_angle3_value(3) = con_angle3_value(0)
Erase angle
 ReDim angle(0) As angle_type
Erase angle_relation.av_no
 ReDim angle_relation.av_no(0) As Av_no_type
Erase angle_value_90.av_no
 ReDim angle_value_90.av_no(0) As Av_no_type
Erase angle_value.av_no
 ReDim angle_value.av_no(0) As Av_no_type
Erase area_of_circle
 ReDim area_of_circle(0) As area_of_circle_type
 con_Area_of_circle(0).data(0) = area_of_circle(0).data(0)
 con_Area_of_circle(0).data(1) = area_of_circle(0).data(0)
 con_Area_of_circle(1) = con_Area_of_circle(0)
 con_Area_of_circle(2) = con_Area_of_circle(0)
 con_Area_of_circle(3) = con_Area_of_circle(0)
Erase Area_of_fan
 ReDim Area_of_fan(0) As area_of_fan_type
 con_Area_of_fan(0).data(0) = Area_of_fan(0).data(0)
 con_Area_of_fan(0).data(1) = Area_of_fan(0).data(0)
 con_Area_of_fan(1) = con_Area_of_fan(0)
 con_Area_of_fan(2) = con_Area_of_fan(0)
 con_Area_of_fan(3) = con_Area_of_fan(0)
Erase area_of_element
 ReDim area_of_element(0) As area_of_element_type
 con_Area_of_element(0).data(0) = area_of_element(0).data(0)
 con_Area_of_element(0).data(1) = area_of_element(0).data(0)
 con_Area_of_element(1) = con_Area_of_element(0)
 con_Area_of_element(2) = con_Area_of_element(0)
 con_Area_of_element(3) = con_Area_of_element(0)
Erase Sides_length_of_triangle
 ReDim Sides_length_of_triangle(0) As sides_length_of_triangle_type
 con_Sides_length_of_triangle(0).data(0) = Sides_length_of_triangle(0).data(0)
 con_Sides_length_of_triangle(0).data(1) = Sides_length_of_triangle(0).data(0)
 con_Sides_length_of_triangle(1) = con_Sides_length_of_triangle(0)
 con_Sides_length_of_triangle(2) = con_Sides_length_of_triangle(0)
 con_Sides_length_of_triangle(3) = con_Sides_length_of_triangle(0)
Erase arc_value
 ReDim arc_value(0) As arc_value_type
 con_arc_value(0).data(0) = arc_value(0).data(0)
 con_arc_value(0).data(1) = arc_value(0).data(0)
 con_arc_value(1) = con_arc_value(0)
 con_arc_value(2) = con_arc_value(0)
 con_arc_value(3) = con_arc_value(0)
Erase arc
 ReDim arc(0) As arc_type
Erase Dangle
Erase Dline1
Erase Ddistance_of_paral_line
 ReDim Ddistance_of_paral_line(0) As distance_of_paral_line_data_type
Erase Ddistance_of_point_line
 ReDim Ddistance_of_point_line(0) As distance_of_point_line_data_type
Erase Ddpoint_pair
 ReDim Ddpoint_pair(0) As Dpoint_pair_type
 con_dpoint_pair(0).data(0) = Ddpoint_pair(0).data(0).data0
 con_dpoint_pair(0).data(1) = Ddpoint_pair(0).data(0).data0
 con_dpoint_pair(1) = con_dpoint_pair(0)
 con_dpoint_pair(2) = con_dpoint_pair(0)
 con_dpoint_pair(3) = con_dpoint_pair(0)
Erase epolygon
 ReDim epolygon(0) As epolygon_type
 con_Epolygon(0).data(0) = epolygon(0).data(0)
 con_Epolygon(0).data(1) = epolygon(0).data(0)
 con_Epolygon(1) = con_Epolygon(0)
 con_Epolygon(2) = con_Epolygon(0)
 con_Epolygon(3) = con_Epolygon(0)
Erase Deangle.av_no
 ReDim Deangle.av_no(0) As Av_no_type
Erase equal_arc
 ReDim equal_arc(0) As equal_arc_type
 con_equal_arc(0).data(0) = equal_arc(0).data(0)
 con_equal_arc(0).data(1) = equal_arc(0).data(0)
 con_equal_arc(1) = con_equal_arc(0)
 con_equal_arc(2) = con_equal_arc(0)
 con_equal_arc(3) = con_equal_arc(0)
Erase equal_side_right_triangle
 ReDim equal_side_right_triangle(0) As one_triangle_type
 con_equal_side_right_triangle(0).data(0) = equal_side_right_triangle(0).data(0)
 con_equal_side_right_triangle(0).data(1) = equal_side_right_triangle(0).data(0)
 con_equal_side_right_triangle(1) = con_equal_side_right_triangle(0)
 con_equal_side_right_triangle(2) = con_equal_side_right_triangle(0)
 con_equal_side_right_triangle(2) = con_equal_side_right_triangle(0)
'Erase equal_area_triangle
' ReDim equal_area_triangle(0) As equal_area_triangle_type
' con_equal_area_triangle(0).data(0) = equal_area_triangle(0).data(0)
' con_equal_area_triangle(0).data(1) = equal_area_triangle(0).data(0)
' con_equal_area_triangle(1) = con_equal_area_triangle(0)
' con_equal_area_triangle(2) = con_equal_area_triangle(0)
' con_equal_area_triangle(3) = con_equal_area_triangle(0)
Erase equation
 ReDim equation(0) As Equation_type
Erase Deline
 ReDim Deline(0) As eline_type
 con_eline(0).data(0).data0 = Deline(0).data(0).data0
 con_eline(0).data(1).data0 = Deline(0).data(0).data0
 con_eline(1) = con_eline(0)
 con_eline(2) = con_eline(0)
 con_eline(3) = con_eline(0)
Erase general_string
 ReDim general_string(0) As general_string_type
 con_general_string(0).data(0) = general_string(0).data(0)
 con_general_string(0).data(1) = general_string(0).data(0)
 con_general_string(1) = con_general_string(0)
 con_general_string(2) = con_general_string(0)
 con_general_string(3) = con_general_string(0)
Erase general_angle_string
 ReDim general_angle_string(0) As general_angle_string_type
 con_general_angle_string(0).data(0) = general_angle_string(0).data(0)
 con_general_angle_string(0).data(1) = general_angle_string(0).data(0)
 con_general_angle_string(1) = con_general_angle_string(0)
 con_general_angle_string(2) = con_general_angle_string(0)
 con_general_angle_string(3) = con_general_angle_string(0)
Erase function_of_angle
 ReDim function_of_angle(0) As function_of_angle_type
 con_function_of_angle(0).data(0) = function_of_angle(0).data(0)
 con_function_of_angle(0).data(1) = function_of_angle(0).data(0)
 con_function_of_angle(1) = con_function_of_angle(0)
 con_function_of_angle(2) = con_function_of_angle(0)
 con_function_of_angle(3) = con_function_of_angle(0)
Erase four_point_on_circle
 ReDim four_point_on_circle(0) As four_point_on_circle_type
 con_Four_point_on_circle(0).data(0) = four_point_on_circle(0).data(0)
 con_Four_point_on_circle(0).data(1) = four_point_on_circle(0).data(0)
 con_Four_point_on_circle(1) = con_Four_point_on_circle(0)
 con_Four_point_on_circle(2) = con_Four_point_on_circle(0)
 con_Four_point_on_circle(3) = con_Four_point_on_circle(0)
Erase item0
 ReDim Preserve item0(0) As item0_type
Erase Dtwo_point_line
 ReDim Dtwo_point_line(0) As line_from_two_point
Erase line_value
 ReDim line_value(0) As line_value_type
 con_line_value(0).data(0) = line_value(0).data(0)
 con_line_value(0).data(1) = line_value(0).data(0)
 con_line_value(1) = con_line_value(0)
 con_line_value(2) = con_line_value(0)
 con_line_value(3) = con_line_value(0)
Erase Dlong_squre
 ReDim Dlong_squre(0) As long_squre_type
 con_long_squre(0).data(0) = Dlong_squre(0).data(0)
 con_long_squre(0).data(1) = Dlong_squre(0).data(0)
 con_long_squre(1) = con_long_squre(0)
 con_long_squre(2) = con_long_squre(0)
 con_long_squre(3) = con_long_squre(0)
Erase Dsqure
 ReDim Dsqure(0) As squre_type
 con_squre(0).data(0) = Dsqure(0).data(0)
 con_squre(0).data(1) = Dsqure(0).data(0)
 con_squre(1) = con_squre(0)
 con_squre(2) = con_squre(0)
 con_squre(3) = con_squre(0)
Erase mid_point_line
 ReDim mid_point_line(0) As mid_point_line_type
 con_mid_point_line(0).data(0) = mid_point_line(0).data(0)
 con_mid_point_line(0).data(1) = mid_point_line(0).data(0)
 con_mid_point_line(1) = con_mid_point_line(0)
 con_mid_point_line(2) = con_mid_point_line(0)
 con_mid_point_line(3) = con_mid_point_line(0)
Erase Dmid_point
 ReDim Dmid_point(0) As mid_point_type
 con_mid_point(0).data(0) = Dmid_point(0).data(0).data0
 con_mid_point(0).data(1) = Dmid_point(0).data(0).data0
 con_mid_point(1) = con_mid_point(0)
 con_mid_point(2) = con_mid_point(0)
 con_mid_point(3) = con_mid_point(0)
Erase new_point
 ReDim new_point(0) As new_point_type
Erase Dparallelogram
 ReDim Dparallelogram(0) As parallelogram_type
Erase Dparal
 ReDim Dparal(0) As paral_type
 con_paral(0).data(0) = Dparal(0).data(0).data0
 con_paral(0).data(1) = Dparal(0).data(0).data0
 con_paral(1) = con_paral(0)
 con_paral(2) = con_paral(0)
 con_paral(3) = con_paral(0)
Erase poly
 ReDim poly(0) As polygon
Erase Dpolygon4
 ReDim Dpolygon4(0) As polygon4_type
 con_parallelogram(0).data(0) = Dpolygon4(0).data(0)
 con_parallelogram(0).data(1) = Dpolygon4(0).data(0)
 con_parallelogram(1) = con_parallelogram(0)
 con_parallelogram(2) = con_parallelogram(0)
 con_parallelogram(3) = con_parallelogram(0)
Erase point_pair_for_similar
 ReDim point_pair_for_similar(0) As point_pair_for_similar_type
Erase relation_from_line_to_triangle
 ReDim relation_from_line_to_triangle(0) As relation_from_line_to_triangle_type
Erase relation_from_triangle_to_line
 ReDim relation_from_triangle_to_line(0) As relation_from_triangle_to_line_type
  'ReDim relation_from_triangle_to_line(0).data(0) As relation_from_triangle_to_line_data_type
Erase Drelation
 ReDim Drelation(0) As relation_type
 con_relation(0).data(0) = Drelation(0).data(0).data0
 con_relation(0).data(1) = Drelation(0).data(0).data0
 con_relation(1) = con_relation(0)
 con_relation(2) = con_relation(0)
 con_relation(3) = con_relation(0)
Erase ratio_of_two_arc
 ReDim ratio_of_two_arc(0) As ratio_of_two_arc_type
 con_ratio_of_two_arc(0).data(0) = ratio_of_two_arc(0).data(0)
 con_ratio_of_two_arc(0).data(1) = ratio_of_two_arc(0).data(0)
 con_ratio_of_two_arc(1) = con_ratio_of_two_arc(0)
 con_ratio_of_two_arc(2) = con_ratio_of_two_arc(0)
 con_ratio_of_two_arc(3) = con_ratio_of_two_arc(0)
Erase rhombus
 ReDim rhombus(0) As rhombus_type
 con_rhombus(0).data(0) = rhombus(0).data(0)
 con_rhombus(0).data(1) = rhombus(0).data(0)
 con_rhombus(1) = con_rhombus(0)
 con_rhombus(2) = con_rhombus(0)
 con_rhombus(3) = con_rhombus(0)
Erase same_three_lines
 ReDim same_three_lines(0) As same_three_lines_type
Erase Sides_length_of_circle
 ReDim Sides_length_of_circle(0) As sides_length_of_circle_type
 con_Sides_length_of_circle(0).data(0) = Sides_length_of_circle(0).data(0)
 con_Sides_length_of_circle(0).data(1) = Sides_length_of_circle(0).data(0)
 con_Sides_length_of_circle(1) = con_Sides_length_of_circle(0)
 con_Sides_length_of_circle(2) = con_Sides_length_of_circle(0)
 con_Sides_length_of_circle(3) = con_Sides_length_of_circle(0)
Erase Dsimilar_triangle
 ReDim Dsimilar_triangle(0) As similar_triangle_type
 con_similar_triangle(0).data(0) = Dsimilar_triangle(0).data(0)
 con_similar_triangle(0).data(1) = Dsimilar_triangle(0).data(0)
 con_similar_triangle(1) = con_similar_triangle(0)
 con_similar_triangle(2) = con_similar_triangle(0)
 con_similar_triangle(3) = con_similar_triangle(0)
Erase Squ_sum
 ReDim Squ_sum(0) As squ_sum_type
Erase string_value
 ReDim string_value(0) As string_value_type
Erase m_tangent_circle
 ReDim tangent_circle(0) As tangent_circle_type
Erase tangent_line
 ReDim tangent_line(0) As tangent_line_type
 con_tangent_line(0).data(0) = tangent_line(0).data(0)
 con_tangent_line(0).data(1) = tangent_line(0).data(0)
 con_tangent_line(1) = con_tangent_line(0)
 con_tangent_line(2) = con_tangent_line(0)
 con_tangent_line(3) = con_tangent_line(0)
Erase T_angle
 ReDim T_angle(0) As total_angle_type
Erase Darea_relation
 ReDim Darea_relation(0) As area_relation_type
Erase tri_function
 ReDim tri_function(0) As tri_function_type
Erase three_angle_value_sum.av_no
 ReDim three_angle_value_sum.av_no(0) As Av_no_type
Erase two_angle_value_sum.av_no
 ReDim two_angle_value_sum.av_no(0) As Av_no_type
Erase two_angle_value_180.av_no
 ReDim two_angle_value_180.av_no(0) As Av_no_type
Erase two_angle_value_90.av_no
 ReDim two_angle_value_90.av_no(0) As Av_no_type
Erase Two_angle_value.av_no
 ReDim Two_angle_value.av_no(0) As Av_no_type
Erase Two_angle_value0.av_no
 ReDim Two_angle_value0.av_no(0) As Av_no_type
Erase two_order_equation
 ReDim two_order_equation(0) As two_order_equation_type
Erase Dtixing
 ReDim Dtixing(0) As tixing_type
Erase Dtotal_equal_triangle
 ReDim Dtotal_equal_triangle(0) As total_equal_triangle_type
 con_total_equal_triangle(0).data(0) = Dtotal_equal_triangle(0).data(0)
 con_total_equal_triangle(0).data(1) = Dtotal_equal_triangle(0).data(0)
 con_total_equal_triangle(1) = con_total_equal_triangle(0)
 con_total_equal_triangle(2) = con_total_equal_triangle(0)
 con_total_equal_triangle(3) = con_total_equal_triangle(0)
Erase pseudo_similar_triangle
 ReDim pseudo_similar_triangle(0) As pseudo_two_triangle_type
Erase pseudo_total_equal_triangle
 ReDim pseudo_total_equal_triangle(0) As pseudo_two_triangle_type
Erase triangle
 ReDim triangle(0) As triangle_type
Erase three_point_on_line
 ReDim three_point_on_line(0) As three_point_on_line_type
 con_Three_point_on_line(0).data(0) = three_point_on_line(0).data(0)
 con_Three_point_on_line(0).data(1) = three_point_on_line(0).data(0)
 con_Three_point_on_line(1) = con_Three_point_on_line(0)
 con_Three_point_on_line(2) = con_Three_point_on_line(0)
 con_Three_point_on_line(3) = con_Three_point_on_line(0)
Erase two_point_conset
 ReDim two_point_conset(0) As two_point_conset_type
Erase two_line_value
 ReDim two_line_value(0) As two_line_value_type
 con_two_line_value(0).data(0) = two_line_value(0).data(0).data0
 con_two_line_value(0).data(1) = two_line_value(0).data(0).data0
 con_two_line_value(1) = con_two_line_value(0)
 con_two_line_value(2) = con_two_line_value(0)
 con_two_line_value(3) = con_two_line_value(0)
Erase line3_value
 ReDim line3_value(0) As line3_value_type
 con_line3_value(0).data(0) = line3_value(0).data(0).data0
 con_line3_value(0).data(1) = line3_value(0).data(0).data0
 con_line3_value(1) = con_line3_value(0)
 con_line3_value(2) = con_line3_value(0)
 con_line3_value(3) = con_line3_value(0)
Erase verti_mid_line
 ReDim verti_mid_line(0) As verti_mid_line_type
 con_verti_mid_line(0).data(0) = verti_mid_line(0).data(0)
 con_verti_mid_line(0).data(1) = verti_mid_line(0).data(0)
 con_verti_mid_line(1) = con_verti_mid_line(0)
 con_verti_mid_line(2) = con_verti_mid_line(0)
 con_verti_mid_line(3) = con_verti_mid_line(0)
Erase Dverti
 ReDim Dverti(0) As verti_type
 con_verti(0).data(0) = Dverti(0).data(0)
 con_verti(0).data(1) = Dverti(0).data(0)
 con_verti(1) = con_verti(0)
 con_verti(2) = con_verti(0)
 con_verti(3) = con_verti(0)
Erase angle_less_angle
 ReDim angle_less_angle(0) As angle_less_angle_type
  'ReDim angle_less_angle(0).data(0) As angle_less_angle_data_type
Erase line_less_line
 ReDim line_less_line(0) As line_less_line_type
  'ReDim line_less_line(0).data(0) As line_less_line_data_type
Erase line_less_line2
 ReDim line_less_line2(0) As line_less_line2_type
  'ReDim line_less_line2(0).data(0) As line_less_line2_data_type
Erase line2_less_line2
 ReDim line2_less_line2(0) As line2_less_line2_type
  'ReDim line2_less_line2(0).data(0) As line2_less_line2_data_type
Erase angle3_value
 ReDim angle3_value(0) As angle3_value_type
  'ReDim Dlong_squre(0).data(0) As long_squre_data_type
  'ReDim Sides_length_of_triangle(0).data(0) As sides_length_of_triangle_data_type
  'ReDim Sides_length_of_circle(0).data(0) As sides_length_of_circle_data_type
  'ReDim Circ(0).data(0) As circle_data_type
'Erase Con_lin
' last_dangle = 0
 ' old_last_dangle = 0
'last_dline1 = 0
 ' old_last_dline1 = 0
' last_squ_sum = 0
  'old_last_squ_sum = 0
Erase verti_mid_line
ReDim verti_mid_line(0) As verti_mid_line_type
 ' last_verti_mid_line(0) = 0
 '  last_verti_mid_line(1) = 0
  '  old_last_verti_mid_line = 0
'******
Erase display_string
'For i% = 0 To 100
'display_record_type(i%) = 0
'display_record_no(i%) = 0
'Next i%
End Sub

Public Sub init_input_cond(n%)
Dim j%
 For j% = 0 To 50
Call C_display_wenti.set_m_condition(n%, empty_char, j%)
Call C_display_wenti.set_m_point_no(n%, 0, j%, False)
Next j%
Call C_display_wenti.set_m_no(0, n%, 0)

End Sub
Public Sub initial_record(record As record_type)
Dim re As record_type
record = re
End Sub

Public Sub init_no_reduce_for_condition()
Dim i%
For i% = 1 To last_conditions.last_cond(1).dpoint_pair_no  '3
    Ddpoint_pair(i%).record_.no_reduce = 0
Next i%
'For i% = 1 To last_conditions.last_cond(1).distance_of_paral_line_no  '3
'    Ddistance_of_paral_line(i%).record_.no_reduce = 0
'Next i%
For i% = 1 To last_conditions.last_cond(1).distance_of_point_line_no  '3
    Ddistance_of_point_line(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).relation_from_line_to_triangle_no  '3
    relation_from_line_to_triangle(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).relation_from_triangle_to_line_no  '3
    relation_from_triangle_to_line(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).triangle_no '4
    triangle(i%).record_.no_reduce = 0
 Next i%
For i% = 1 To last_conditions.last_cond(1).rtriangle_no '4
    Rtriangle(i%).record_.no_reduce = 0
 Next i%
For i% = 1 To last_conditions.last_cond(1).area_relation_no  '4
    Darea_relation(i%).record_.no_reduce = 0
 Next i%
For i% = 1 To last_conditions.last_cond(1).mid_point_line_no  '7
      mid_point_line(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).eline_no '8
     Deline(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).four_point_on_circle_no '9
    four_point_on_circle(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).three_point_on_circle_no '9
    three_point_on_circle(i%).record_.no_reduce = 0
Next i%

For i% = 1 To last_conditions.last_cond(1).mid_point_no '12
     Dmid_point(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).paral_no '13
     Dparal(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).parallelogram_no
     Dparallelogram(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).relation_no '16
     Drelation(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).similar_triangle_no  '17
     Dsimilar_triangle(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).total_equal_triangle_no  '18
     Dtotal_equal_triangle(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).pseudo_total_equal_triangle_no  '18
     pseudo_total_equal_triangle(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).pseudo_similar_triangle_no  '18
     pseudo_similar_triangle(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).three_point_on_line_no '20
     three_point_on_line(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).two_point_conset_no '20
     two_point_conset(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).two_point_conset_no '20
     two_point_conset(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).two_line_value_no
     two_line_value(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).line3_value_no
    line3_value(i%).record_.no_reduce = 0
Next i%
 For i% = 1 To last_conditions.last_cond(1).verti_no '24
     Dverti(i%).record_.no_reduce = 0
Next i%
 'For i% = 1 To last_conditions.last_cond(1).vector_no '24
'     Dvector(i%).record_.no_reduce = 0
'Next i%
For i% = 1 To last_conditions.last_cond(1).arc_value_no '25
    arc_value(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).equal_arc_no '26
     equal_arc(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).ratio_of_two_arc_no '27
     ratio_of_two_arc(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).angle_less_angle_no '28
     angle_less_angle(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).line_less_line_no '29
     line_less_line(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).line_less_line2_no '30
     line_less_line2(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).line2_less_line2_no '31
     line2_less_line2(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).line_value_no '33
     line_value(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).tangent_line_no  '34
     tangent_line(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).tangent_circle_no '34
     tangent_circle(i%).record_.no_reduce = 0
Next i%
'For i% = 1 To last_conditions.last_cond(1).equal_area_triangle_no '35
 '   equal_area_triangle(i%).record_.no_reduce = 0
'Next i%
For i% = 1 To last_conditions.last_cond(1).general_string_no '36
     general_string(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).general_angle_string_no  '37
     general_angle_string(i%).record_.no_reduce = 0
Next i%
'For i% = 1 To last_conditions.last_cond(1).equal_side_tixing_no '38
'     Dequal_side_tixing(i%).record_.no_reduce = 0
'Next i%
For i% = 1 To last_conditions.last_cond(1).equal_side_triangle_no '38
     equal_side_triangle(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).equation_no '38
     equation(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).epolygon_no '39
     epolygon(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).tixing_no '40
     Dtixing(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).rhombus_no '41
    rhombus(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).long_squre_no '42
     Dlong_squre(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).squre_no '42
     Dsqure(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).area_of_element_no '43
     area_of_element(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).area_of_circle_no '44
     area_of_circle(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).area_of_fan_no '46
     Area_of_fan(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).sides_length_of_triangle_no '47
     Sides_length_of_triangle(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).sides_length_of_circle_no '48
     Sides_length_of_circle(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).verti_mid_line_no  '49
     verti_mid_line(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).v_line_value_no  '49
     V_line_value(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).V_two_line_time_no  '49
     V_two_line_time_value(i%).record_.no_reduce = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).squ_sum_no '50
     Squ_sum(i%).record_.no_reduce = 0
Next i%
End Sub

Public Sub init_condition_no(cond_no As conditions_data_type)
'cond_no.total_condition = 0
cond_no.pass_word_for_teacher = "0000*"
cond_no.init_v_line_no = 0
cond_no.value_string_no = 0
cond_no.aid_point_data1_no = 0
cond_no.aid_point_data2_no = 0
cond_no.aid_point_data3_no = 0
cond_no.branch_data_no = 0
cond_no.new_point_no = 0
cond_no.note_space_no = 0
cond_no.unkown_element_no = 0
cond_no.aid_point_no = 0 '15
cond_no.four_sides_fig_no = 0
cond_no.angle3_value_no = 0 '32
cond_no.angle_less_angle_no = 0 '28
cond_no.angle_no = 0 '1
cond_no.angle_relation_no = 0
cond_no.angle_value_90_no = 0
cond_no.angle_value_no = 0
cond_no.area_of_circle_no = 0 '44
cond_no.area_of_fan_no = 0 '46
cond_no.area_of_element_no = 0 '43
cond_no.sides_length_of_triangle_no = 0 '47
cond_no.arc_value_no = 0 '25
cond_no.arc_no = 0
cond_no.change_picture_type = 0
cond_no.change_picture_step = 0
cond_no.con_line_no = 0
cond_no.dangle_no = 0
cond_no.dline1_no = 0
cond_no.distance_of_paral_line_no = 0
cond_no.distance_of_point_line_no = 0
cond_no.dpoint_pair_no = 0 '3
cond_no.epolygon_no = 0 '40
cond_no.eangle_no = 0
cond_no.equal_3angle_no = 0
cond_no.equal_arc_no = 0 '26
cond_no.equal_side_right_triangle_no = 0
'cond_no.equal_area_triangle_no = 0 '35
cond_no.polygon4_no = 0
'cond_no.point_pair_for_simlilar_no = 0
'cond_no.equal_side_tixing_no = 0
cond_no.equal_side_triangle_no = 0
cond_no.equation_no = 0
cond_no.eline_no = 0 '8
cond_no.general_string_no = 0  '36
cond_no.general_angle_string_no = 0 '37
cond_no.four_point_on_circle_no = 0 '9
cond_no.three_point_on_circle_no = 0 '9
cond_no.function_of_angle_no = 0
cond_no.general_string_combine_no = 0
cond_no.item0_no = 0
cond_no.last_angle3_value_combine = 0
cond_no.last_general_string_combine = 0
cond_no.length_of_polygon_no = 0
cond_no.line_from_two_point_no = 0
cond_no.line2_less_line2_no = 0 '31
cond_no.line3_value_no = 0
cond_no.line_less_line2_no = 0 '30
cond_no.line_less_line_no = 0 '29
cond_no.line_no = 0 '11
cond_no.line_value_no = 0 '33
cond_no.long_squre_no = 0 '42
cond_no.squre_no = 0
cond_no.mid_point_line_no = 0 '7
cond_no.mid_point_no = 0 '12
cond_no.new_point_no = 0
cond_no.set_branch = False
cond_no.new_midpoint_no = 0
cond_no.parallelogram_no = 0 '14
cond_no.paral_no = 0 '13
cond_no.point_no = 0
cond_no.poly_no = 0
cond_no.point_pair_for_similar_no = 0
cond_no.pre_add_condition_no = 0
cond_no.pseudo_dpoint_pair_no = 0
cond_no.pseudo_midpoint_no = 0
cond_no.pseudo_relation_no = 0
cond_no.pseudo_line3_value_no = 0
cond_no.pseudo_eline_no = 0
cond_no.relation_from_line_to_triangle_no = 0 '3
cond_no.relation_from_triangle_to_line_no = 0 '3
cond_no.relation_no = 0 '16
cond_no.v_relation_no = 0 '16
cond_no.relation_on_line_no = 0 '16
cond_no.relation_string_no = 0
cond_no.ratio_of_two_arc_no = 0 '27
cond_no.rhombus_no = 0 '41
cond_no.right_angle_for_Pd_no = 0
cond_no.rtriangle_no = 0
cond_no.same_three_lines_no = 0
cond_no.sides_length_of_circle_no = 0 '48
cond_no.similar_triangle_no = 0 '17
cond_no.string_value_no = 0
cond_no.squ_sum_no = 0 '50
cond_no.tangent_circle_no = 0
cond_no.tangent_line_no = 0 '34
cond_no.three_angle_value0_no = 0
cond_no.total_angle_no = 0 '1
cond_no.total_condition = 0
cond_no.trajectory_no = 0
cond_no.area_relation_no = 0 '4
cond_no.tri_function_no = 0
cond_no.three_angle_value_sum_no = 0
cond_no.two_area_of_element_value_no = 0
cond_no.two_angle_value_sum_no = 0
cond_no.two_angle_value_180_no = 0
cond_no.two_angle_value_90_no = 0
cond_no.two_angle_value_no = 0
cond_no.two_angle_value0_no = 0
cond_no.two_order_eqution_no = 0
cond_no.tixing_no = 0 '39
cond_no.total_equal_triangle_no = 0 '18
cond_no.pseudo_total_equal_triangle_no = 0 '18
cond_no.pseudo_similar_triangle_no = 0 '18
cond_no.triangle_no = 0 '19
cond_no.three_point_on_line_no = 0 '20
cond_no.two_point_conset_no = 0 '20
cond_no.two_line_value_no = 0 '21
cond_no.verti_mid_line_no = 0 '49
cond_no.verti_no = 0 '24
cond_no.last_view_point_no = 0
cond_no.last_dot_line = 0
cond_no.v_line_value_no = 0
cond_no.V_two_line_time_no = 0
End Sub

Public Sub init_inform()
MDIForm1.angle_inform.Enabled = False
MDIForm1.two_angle.Enabled = False
MDIForm1.angle_relation = False
MDIForm1.three_angle.Enabled = False
MDIForm1.sum_two_angle_right.Enabled = False
MDIForm1.sum_two_angle_pi.Enabled = False
MDIForm1.eangle.Enabled = False
MDIForm1.yizhiA.Enabled = False
MDIForm1.right_angle.Enabled = False
MDIForm1.inform_line.Enabled = False
MDIForm1.paral.Enabled = False
MDIForm1.verti.Enabled = False
MDIForm1.three_point_on_line.Enabled = False
MDIForm1.inform_segment.Enabled = False
MDIForm1.length_of_segment.Enabled = False
MDIForm1.two_line_value.Enabled = False
MDIForm1.eline.Enabled = False
MDIForm1.relation.Enabled = False
MDIForm1.re_line.Enabled = False
MDIForm1.inform_circle.Enabled = False
MDIForm1.four_point_on_circle.Enabled = False
MDIForm1.area_of_circle.Enabled = False
MDIForm1.inform_triangle.Enabled = False
MDIForm1.total_equal_triangle.Enabled = False
MDIForm1.similar_triangle.Enabled = False
MDIForm1.area_of_triangle.Enabled = False
MDIForm1.infrom_polygon.Enabled = False
MDIForm1.sp_four_sides.Enabled = False
MDIForm1.area_of_polygon.Enabled = False
MDIForm1.epolygon.Enabled = False
End Sub
Public Sub set_initial_condition(num As Integer, ByVal no_reduce As Byte, ByVal input_type As Boolean)
Dim i%, j%, c%, m%, n%, k%, tp%, tn_%, poly_no%
Dim triA(2) As Integer '输入条件
Dim it(3) As Integer
Dim t_p(3) As Integer
Dim tl(2) As Integer
Dim tn(7) As Integer
Dim ang(2) As Integer
Dim value(2) As String
Dim cir(1) As Integer
Dim cond_ty As Byte
Dim pol As polygon
Dim p(1) As Integer
Dim c_data As condition_data_type
Dim temp_record As total_record_type
 temp_record.record_data.data0.condition_data.condition_no = 0
 temp_record.record_.display_no = -num
 If is_old_conclusion(num) Then
     Exit Sub
 End If
Call control_menu(True) '恢复激活菜单
If input_type Then
   If C_display_wenti.m_no(num) >= 23 Then
    Call init_draw_data
   End If
End If
'if draw_statue=
 MDIForm1.Inputcond.Enabled = True
 MDIForm1.conclusion.Enabled = True
 If num = 0 Then
 last_conditions.last_cond(1).branch_data_no = 0
 Erase branch_data
 ReDim Preserve branch_data(0) As branch_data_type
 last_conditions.last_cond(1).set_branch = True
 End If
 Call set_point_from_wenti_no(num, C_display_wenti.m_no(num))
Select Case C_display_wenti.m_no(num)
'Case -56
'Call set_initial_condition_56(num, temp_record)
'-70 直线□□上取一定点□
'-69 直线□□上[任取一/取一定]点□
'-63 □是直线□□与⊙□□□的一个交点
'-62 过□平行□□的直线交⊙□□□于□
'-57 在⊙□□□上取一点□使得□□＝!_~
'-56 ∠□□□的平分线交⊙□□□于□
Case -54, -53, -25
'-54 □□的垂直平分线交□□于□
'-53 □□的垂直平分线交⊙□[down\\(_)]于□
'-25 □□的垂直平分线交⊙□□□于□
Call set_initial_condition_54_53_25(num, temp_record)
Case -52, -51, -56
'-52 ∠□□□的平分线交⊙□[down\\(_)]于□
'-51 ∠□□□的平分线交□□于□
Call set_initial_condition_52_51_56(num, temp_record)
Case -50
'-50 □□是∠□□□的平分线
Call set_initial_condition_50(num, temp_record)
Case -48, -47
'-48 △□□□的周长=!_~
'-47 △□□□的面积=!_~
Call set_initial_condition_48_47(num, temp_record)
Case -46, -45
'-46 四边形□□□□的周长=!_~
'-45 四边形□□□□的面积=!_~
Call set_initial_condition_46_45(num, temp_record)
Case -43, -42
'-43 在□□上取一点□使得□□＝!_~
'-42 在⊙□[down\\(_)]上取一点□使得□□＝!_~
Call set_initial_condition_43_42(num, temp_record)
Case -41
'-41 ∠□□□/∠□□□=!_~
Call set_initial_condition_41(num, temp_record)
Case -40
'-40 □□/□□＝□□/□□
Call set_initial_condition_40(num, temp_record)
Case -39
'-39 ∠□□□=∠□□□+∠□□□
Call set_initial_condition_39(num, temp_record)
Case -38 '∠□□□+∠□□□=!_~°
'-38 ∠□□□+∠□□□=!_~°
Call set_initial_condition_38(num, temp_record)
Case -37 '△□□□≌△□□□
'-37 △□□□≌△□□□
Call set_initial_condition_37(num, temp_record)
Case -36 '△□□□∽△□□□
'-36 △□□□∽△□□□
Call set_initial_condition_36(num, temp_record)
Case -35 '□□=□□+□□
'-35 □□＝□□+□□
Call set_initial_condition_35(num, temp_record)
Case -34 '□□+□□=!_~
'-34 □□+□□＝!_~
Call set_initial_condition_34(num, temp_record)
Case -33, -44
'-33 过□作⊙□[down\\(_)]的切线□□
'-44 过□作⊙□□□的切线□□
Call set_initial_condition_33_44(num, temp_record)
Case -32, -3, -29, -28, -26, -27
'-3 与⊙□[down\\(_)]相切于点□的切线交⊙□[down\\(_)]
'-32 与⊙□[down\\(_)]相切于点□的切线交直线□□于□
'-29 与⊙□□□相切于□的切线交⊙□[down\\(_)]于□
'-28 与⊙□□□相切于□的切线交⊙□□□于□
'-27 与⊙□□□相切于□的切线交直线□□于□
'-26 与⊙□[down\\(_)]相切于□的切线交⊙□□□于□
Call set_initial_condition_32_3(num, temp_record)
Case -31
'-31 在□□上取一点□使得□□＝□□
Call set_initial_condition_31(num, temp_record)
Case -30, -58
'-58 在⊙□□□上取一点□使得□□＝□□
'-30 在⊙□[down\\(_)]上取一点□使得□□＝□□
Call set_initial_condition_30_58(num, temp_record)
Case -24
'-24 弧□□＝弧□□
Call set_initial_condition_24(num, temp_record)
Case -23, -22
'-22 过□点平行□□的直线交□□于□
'-23过□点垂直□□的直线交□□于□
Call set_initial_condition_23_22(num, temp_record)
Case -20
'-20 任意△□□□
'Call set_initial_condition_20(num)
'depend_no(num) = 1
'inpcond(-21) = △□□□是直角三角形
'inpcond(-20) = 任意△□□□
'Call draw_picture_20(num, no_reduce)
Case -21
'-21 △□□□是直角三角形
Call set_initial_condition_21(num, temp_record)

Case -19
'-19 任意四边形□□□□
'Call draw_picture_19(num, no_reduce)
Case -17, -18
'-17 △□□□是等腰直角三角形
'-18 △□□□是等腰三角形
Call set_initial_condition_17_18(num, temp_record)
Case -16, -12, -9, -8
'-8 □□□□□□是正六边形
'-9 □□□□□是正五边形
'-12 □□□□是正方形
'-16 △□□□是等边三角形
 If C_display_wenti.m_no(num) = -16 Then
  j% = 3
 ElseIf C_display_wenti.m_no(num) = -12 Then
  j% = 4
 ElseIf C_display_wenti.m_no(num) = -9 Then
  j% = 5
 ElseIf C_display_wenti.m_no(num) = -8 Then
  j% = 6
 End If
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
If C_display_wenti.m_no(num) <> -12 Then
If last_conditions.last_cond(1).poly_no = last_conditions.last_cond(2).poly_no Then
ReDim Preserve poly(last_conditions.last_cond(2).poly_no + 10) As polygon
last_conditions.last_cond(2).poly_no = last_conditions.last_cond(2).poly_no + 10
End If
last_conditions.last_cond(1).poly_no = last_conditions.last_cond(1).poly_no + 1
poly_no% = last_conditions.last_cond(1).poly_no
poly(poly_no%) = polygon_data_0
poly(poly_no%).total_v = j%
For i% = 0 To j% - 1
poly(poly_no%).v(i%) = _
            C_display_wenti.m_point_no(num, i%)
Next i%
'****************
 pol.total_v = j%
 pol.v(0) = C_display_wenti.m_point_no(num, 0)
 pol.v(1) = C_display_wenti.m_point_no(num, 1)
 For i% = 2 To j% - 1
 pol.v(i%) = C_display_wenti.m_point_no(num, i%)
'Call set_point_depend_element(pol.v(i%), point_, pol.v(0), point_, pol.v(1))
 Next i%
tn_% = 0
Call set_Epolygon(pol, temp_record, tn_%, 0, 0, cond_ty, True)
Call C_display_wenti.set_data_condition(num, cond_ty, tn_%)
Else 'If c_display_wenti.m_no(num) = -12 Then
 n% = 0
 Call line_number(C_display_wenti.m_point_no(num, 0), _
       C_display_wenti.m_point_no(num, 1), pointapi0, pointapi0, _
        depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
        depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
        condition, condition_color, 1, 0)
 Call line_number(C_display_wenti.m_point_no(num, 1), _
       C_display_wenti.m_point_no(num, 2), pointapi0, pointapi0, _
       depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
       depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
       condition, condition_color, 1, 0)
 Call line_number(C_display_wenti.m_point_no(num, 2), _
       C_display_wenti.m_point_no(num, 3), pointapi0, pointapi0, _
       depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
       depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
       condition, condition_color, 1, 0)
 Call line_number(C_display_wenti.m_point_no(num, 3), _
       C_display_wenti.m_point_no(num, 0), pointapi0, pointapi0, _
       depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
       depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
       condition, condition_color, 1, 0)
 Call set_squre(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
     C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), 0, temp_record, 0, True)
End If
'*******************************************************
Case -15
'-15 □□□□是梯形
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
  Call set_tixing(C_display_wenti.m_point_no(num, 0), _
         C_display_wenti.m_point_no(num, 1), _
          C_display_wenti.m_point_no(num, 2), _
            C_display_wenti.m_point_no(num, 3), tixing_, temp_record, 0, no_reduce)
  Call arrange_four_point_for_input_order(C_display_wenti.m_point_no(num, 0), _
          C_display_wenti.m_point_no(num, 1), _
          C_display_wenti.m_point_no(num, 2), _
           C_display_wenti.m_point_no(num, 3), t_p(0), t_p(1), t_p(2), t_p(3))
  t_line(0) = line_number(t_p(0), t_p(1), pointapi0, pointapi0, _
                          depend_condition(point_, t_p(0)), depend_condition(point_, t_p(1)), _
                          condition, condition_color, 1, 0)
  t_line(1) = line_number(t_p(2), t_p(3), pointapi0, pointapi0, _
                          depend_condition(point_, t_p(2)), depend_condition(point_, t_p(3)), _
                          condition, condition_color, 1, 0)
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
tn_% = 0
Call set_dparal(t_line(0), t_line(1), temp_record, tn_%, 0, True)
Call C_display_wenti.set_m_condition_data(num, paral_, tn_%)
'******************************************************************
Case -13, -11, -14, -10
'-10 □□□□是菱形
'-11 □□□□是平行四边形
'-13 □□□□是长方形
'-14 □□□□是等腰梯形
Call set_initial_condition_14_13_11_10(num, temp_record)

Case -7
'-7 □□/□□=!_~
value(0) = initial_string(number_string(C_display_wenti.m_point_no(num, 4))) 'initial_string(cond_to_string(num, 4, 18, 0))
      If InStr(1, value(0), ".", 0) > 0 Then
        th_chose(-5).chose = 2
      End If
'Call read_number_from_wenti(num, 4, 0, 0, value1)
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
tn_% = 0
Call set_Drelation(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), 0, 0, 0, 0, 0, 0, value(0), _
   temp_record, tn_%, 0, 0, 0, 0, True)
   Call C_display_wenti.set_m_condition_data(num, relation_, tn_%)
Case -6
'-6 □□=!_~
 value(0) = initial_string(number_string(C_display_wenti.m_point_no(num, 2))) 'initial_string(cond_to_string(num, 2, 18, 0))
      If InStr(1, value(0), ".", 0) > 0 Then
        th_chose(-5).chose = 2
      End If
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
tn_% = 0
Call set_line_value(C_display_wenti.m_point_no(num, 0), _
    C_display_wenti.m_point_no(num, 1), _
      value(0), 0, 0, 0, temp_record.record_data, tn_%, 0, True)
       Call C_display_wenti.set_m_condition_data(num, line_value_, tn_%)
Case -5
'-5 ∠□□□=!_~°
Call set_initial_condition_5(num, temp_record)
Case -4
'-4 ∠□□□=∠□□□
Call set_initial_condition_4(num, temp_record)

Case 0     '画点
'depend_no(num) = 1
'Call draw_picture0(num, no_reduce)
Case -2
'-2 作⊙□[down\\(_)]和⊙□[down\\(_)]的公切线□□
'-60 作⊙□□□和⊙□□□的公切线□□
'-59 作⊙□□□和⊙□[down\\(_)]的公切线□□
Call set_initial_condition_2(num, temp_record)
Case -1
Call set_initial_condition_1(num, temp_record)
'-1 □□＝□□
 Case 1 '线上任取一点
Call set_initial_condition1(num, temp_record)
Case 2, 3  ' 平行垂直上任取一点'***
Call set_initial_condition2_3(num, temp_record)
Case 4 ' 垂直平分
'新点应加在最后
Call set_initial_condition4(num, temp_record)
Case 5, 15 '中点
Call set_initial_condition5_15(num, temp_record)
Case 6 '定比分点
Call set_initial_condition6(num, temp_record)
Case 7, -61 ' 圆上任取一点
'-61 ⊙□□□上任取一点□
Call set_initial_condition7_61(num, temp_record)
Case 8, -71
Call set_initial_condition8_71(num, temp_record)
Case 9 '两直线相交
Call set_initial_condition9(num, temp_record)
Case 10, 16, -68  '直线与圆已交于一点求另一交点
'10 过□平行□□的直线交⊙□[down\\(_)]于□
'-68 过□垂直□□的直线交⊙□□□于□
'16  过□垂直□□的直线交⊙□[down\\(_)]于□
Call set_initial_condition10_16(num, temp_record)
Case 11, -63
Call set_initial_condition11_63(num, temp_record)
'depend_no(num) = 1
'Call draw_picture11(num, no_reduce)
Case 12 '两圆相切
'12 ⊙□[down\\(_)]和⊙□[down\\(_)]相切于点□
'-65 ⊙□□□和⊙□□□相切于点□
'-64 ⊙□□□和⊙□[down\\(_)]相切于点□
Call set_initial_condition12(num, temp_record)
Case 13, -67, -66 '两圆的一个交点
'13 □是⊙□[down\\(_)]和⊙□[down\\(_)]一个交点
'-67 □是⊙□□□和⊙□□□一个交点
'-66 □是⊙□□□和⊙□[down\\(_)]一个交点
Call set_initial_condition13(num, temp_record)
Case 14  '过□作直线□□的垂线垂足为□
Call set_initial_condition14(num, temp_record)
'Case 17, 15, 16 '平行   垂直
'17过□垂直□□的直线交⊙□□□于□
'Call draw_picture15_16_17(num, no_reduce)
Case 18, 19, 20, 21 '□是△□□□的重心
'18 □是△□□□的重心
'19 □是△□□□的外接圆的圆心
'20 □是△□□□的垂心
'21 □是△□□□的内切圆的圆心
If C_display_wenti.m_no(num) = 18 Then
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
tp% = 0
tl(0) = line_number(C_display_wenti.m_point_no(num, 2), _
                    C_display_wenti.m_point_no(num, 3), _
                    pointapi0, pointapi0, _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                    condition, condition_color, 1, 0)
tl(1) = line_number(C_display_wenti.m_point_no(num, 1), _
                    C_display_wenti.m_point_no(num, 0), _
                    pointapi0, pointapi0, _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                    condition, condition_color, 1, 0)
tp% = is_line_line_intersect(tl(0), tl(1), 0, 0, True)
If tp% > 0 Then
Call set_mid_point(C_display_wenti.m_point_no(num, 2), tp%, _
      C_display_wenti.m_point_no(num, 3), 0, 0, 0, 0, 0, temp_record, tn_%, 0, 0, 0, 1)
Call set_Drelation(C_display_wenti.m_point_no(num, 1), _
      C_display_wenti.m_point_no(num, 0), _
       C_display_wenti.m_point_no(num, 0), tp%, _
        0, 0, 0, 0, 0, 0, "2", temp_record, 0, 0, 0, 0, 0, False)
End If
'       Call C_display_wenti.set_m_condition_data(num,midpoint_, tn_%)
   temp_record.record_data.data0.condition_data.condition_no = 0 'record0
'   Call add_point_to_line(tp%, tl(0), 0, False, False, 0, c_data)
tp% = 0
tl(0) = line_number(C_display_wenti.m_point_no(num, 1), _
                    C_display_wenti.m_point_no(num, 3), _
                    pointapi0, pointapi0, _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                    condition, condition_color, 1, 0)
tl(1) = line_number(C_display_wenti.m_point_no(num, 2), _
                    C_display_wenti.m_point_no(num, 0), _
                    pointapi0, pointapi0, _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                    condition, condition_color, 1, 0)
tp% = is_line_line_intersect(tl(0), tl(1), 0, 0, True)
If tp% > 0 Then
Call set_mid_point(C_display_wenti.m_point_no(num, 3), tp%, _
      C_display_wenti.m_point_no(num, 1), 0, 0, 0, 0, 0, temp_record, 0, 0, 0, 0, 1)
      Call C_display_wenti.set_m_point_no(num, tp%, 5, True)
Call set_Drelation(C_display_wenti.m_point_no(num, 2), _
      C_display_wenti.m_point_no(num, 0), _
       C_display_wenti.m_point_no(num, 0), tp%, _
        0, 0, 0, 0, 0, 0, "2", temp_record, 0, 0, 0, 0, 0, False)
End If
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
tp% = 0
tl(0) = line_number(C_display_wenti.m_point_no(num, 1), _
                    C_display_wenti.m_point_no(num, 2), _
                    pointapi0, pointapi0, _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                    condition, condition_color, 1, 0)
tl(1) = line_number(C_display_wenti.m_point_no(num, 3), _
                    C_display_wenti.m_point_no(num, 0), _
                    pointapi0, pointapi0, _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                    condition, condition_color, 1, 0)
tp% = is_line_line_intersect(tl(0), tl(1), 0, 0, True)
If tp% > 0 Then
Call set_mid_point(C_display_wenti.m_point_no(num, 1), tp%, _
      C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, temp_record, 0, 0, 0, 0, 1)
      Call C_display_wenti.set_m_point_no(num, tp%, 6, True)
Call set_Drelation(C_display_wenti.m_point_no(num, 2), _
      C_display_wenti.m_point_no(num, 0), _
       C_display_wenti.m_point_no(num, 0), tp%, _
        0, 0, 0, 0, 0, 0, "2", temp_record, 0, 0, 0, 0, 0, False)
End If
'Call set_point_generate_by_line(C_display_wenti.m_point_no(num,num,6), tl(0), 0)
'Call set_Drelation(C_display_wenti.m_point_no(num,num,1), C_display_wenti.m_point_no(num,num,0), _
       C_display_wenti.m_point_no(num,num,0), C_display_wenti.m_point_no(num,num,4), 0, 0, 0, 0, _
        0, 0, "2", temp_record, 0, 0, 0, 0, 0)
'tl(0) = line_number(C_display_wenti.m_point_no(num,num,4), _
                   C_display_wenti.m_point_no(num,num,1), condition, False, 0)
'Call set_point_generate_by_line(C_display_wenti.m_point_no(num,num,0), tl(0), 0)
'Call set_Drelation(C_display_wenti.m_point_no(num,num,2), C_display_wenti.m_point_no(num,num,3), _
       C_display_wenti.m_point_no(num,num,5), C_display_wenti.m_point_no(num,num,6), 0, 0, 0, _
         0, 0, 0, "2", temp_record, 0, 0, 0, 0, 0)
'Call set_dparal(line_number(C_display_wenti.m_point_no(num,num,2), _
                     C_display_wenti.m_point_no(num,num,3), condition, False, 0), _
                line_number(C_display_wenti.m_point_no(num,num,5), _
                     C_display_wenti.m_point_no(num,num,6), condition, False, 0), _
        temp_record, 0, 0)
'Call set_Drelation(C_display_wenti.m_point_no(num,num,2), C_display_wenti.m_point_no(num,num,0), _
       C_display_wenti.m_point_no(num,num,0), C_display_wenti.m_point_no(num,num,5), 0, 0, 0, 0, _
        0, 0, "2", temp_record, 0, 0, 0, 0, 0)
'tl(0) = line_number(C_display_wenti.m_point_no(num,num,2), _
                C_display_wenti.m_point_no(num,num,5), condition, False, 0)
'Call set_point_generate_by_line(C_display_wenti.m_point_no(num,num,0), tl(0), 0)
'Call set_Drelation(C_display_wenti.m_point_no(num,num,1), C_display_wenti.m_point_no(num,num,3), _
       C_display_wenti.m_point_no(num,num,4), C_display_wenti.m_point_no(num,num,6), 0, 0, 0, _
         0, 0, 0, "2", temp_record, 0, 0, 0, 0, 0)
'Call set_dparal(line_number(C_display_wenti.m_point_no(num,num,1), _
                            C_display_wenti.m_point_no(num,num,3), condition, False, 0), _
                line_number(C_display_wenti.m_point_no(num,num,4), _
                           C_display_wenti.m_point_no(num,num,6), condition, False, 0), _
                             temp_record, 0, 0)
'Call set_Drelation(C_display_wenti.m_point_no(num,num,3), C_display_wenti.m_point_no(num,num,0), _
       C_display_wenti.m_point_no(num,num,0), C_display_wenti.m_point_no(num,num,6), 0, 0, 0, 0, _
        0, 0, "2", temp_record, 0, 0, 0, 0, 0)
'tl(0) = line_number(C_display_wenti.m_point_no(num,num,6), _
              C_display_wenti.m_point_no(num,num,0), condition, False, 0)
'Call add_depend_point_for_line(tl(0), C_display_wenti.m_point_no(num,num,6), _
          C_display_wenti.m_point_no(num,num,0))
'Call set_point_generate_by_line(C_display_wenti.m_point_no(num,num,6), tl(0), 0)
'Call set_Drelation(C_display_wenti.m_point_no(num,num,2), C_display_wenti.m_point_no(num,num,1), _
       C_display_wenti.m_point_no(num,num,5), C_display_wenti.m_point_no(num,num,4), 0, 0, 0, _
         0, 0, 0, "2", temp_record, 0, 0, 0, 0, 0)
'Call set_dparal(line_number(C_display_wenti.m_point_no(num,num,2), C_display_wenti.m_point_no(num,num,1), _
       condition, False, 0), line_number(C_display_wenti.m_point_no(num,num,5), C_display_wenti.m_point_no(num,num,4), _
        condition, False, 0), _
        temp_record, 0, 0)
'triA(0) = triangle_number(C_display_wenti.m_point_no(num,num,0), _
   C_display_wenti.m_point_no(num,num,1), C_display_wenti.m_point_no(num,num,2), _
    0, 0, 0, 0, 0, 0, 0)
'triA(1) = triangle_number(C_display_wenti.m_point_no(num,num,0), _
   C_display_wenti.m_point_no(num,num,2), C_display_wenti.m_point_no(num,num,3), _
    0, 0, 0, 0, 0, 0, 0)
'triA(2) = triangle_number(C_display_wenti.m_point_no(num,num,0), _
   C_display_wenti.m_point_no(num,num,3), C_display_wenti.m_point_no(num,num,1), _
    0, 0, 0, 0, 0, 0, 0)
'temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
'Call set_equal_area_triangle(triA(0), triA(1), temp_record, 0, 0, 0)
'temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
'Call set_equal_area_triangle(triA(1), triA(2), temp_record, 0, 0, 0)
'temp_record.record_data.data0.condition_data.condition_no = 0 'record0
'Call set_equal_area_triangle(triA(2), triA(0), temp_record, 0, 0, 0)
ElseIf C_display_wenti.m_no(num) = 19 Then '外心
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
Call set_equal_dline(C_display_wenti.m_point_no(num, 1), _
             C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 2), _
             C_display_wenti.m_point_no(num, 0), 0, 0, 0, 0, 0, 0, 0, _
              temp_record, 0, 0, 0, 0, no_reduce, False)
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
Call set_equal_dline(C_display_wenti.m_point_no(num, 2), _
             C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 3), _
             C_display_wenti.m_point_no(num, 0), 0, 0, 0, 0, 0, 0, 0, _
              temp_record, 0, 0, 0, 0, no_reduce, False)
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
 Call set_equal_dline(C_display_wenti.m_point_no(num, 1), _
             C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 3), _
             C_display_wenti.m_point_no(num, 0), 0, 0, 0, 0, 0, 0, 0, _
             temp_record, 0, 0, 0, 0, no_reduce, False)
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
Call m_circle_number(1, C_display_wenti.m_point_no(num, 0), _
                       m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.coordinate, _
                        C_display_wenti.m_point_no(num, 1), _
                         C_display_wenti.m_point_no(num, 2), _
                          C_display_wenti.m_point_no(num, 3), _
                           0, 0, 0, 1, 0, condition, condition_color, True)
ElseIf C_display_wenti.m_no(num) = 20 Then '垂心
   t_line(0) = line_number(C_display_wenti.m_point_no(num, 0), _
                           C_display_wenti.m_point_no(num, 1), _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                           condition, condition_color, 1, 0)
   t_line(1) = line_number(C_display_wenti.m_point_no(num, 2), _
                           C_display_wenti.m_point_no(num, 3), _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                           condition, condition_color, 1, 0)
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
   Call set_dverti(t_line(0), t_line(1), temp_record, 0, 0, True)
   t_line(0) = line_number(C_display_wenti.m_point_no(num, 0), _
                           C_display_wenti.m_point_no(num, 2), _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                           condition, condition_color, 1, 0)
    t_line(1) = line_number(C_display_wenti.m_point_no(num, 3), _
                           C_display_wenti.m_point_no(num, 1), _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                           condition, condition_color, 1, 0)
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
   Call set_dverti(t_line(0), t_line(1), temp_record, 0, 0, True)
   t_line(0) = line_number(C_display_wenti.m_point_no(num, 1), _
                           C_display_wenti.m_point_no(num, 2), _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                           condition, condition_color, 1, 0)
   t_line(1) = line_number(C_display_wenti.m_point_no(num, 0), _
                           C_display_wenti.m_point_no(num, 3), _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                           condition, condition_color, 1, 0)
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
   Call set_dverti(t_line(0), t_line(1), temp_record, 0, 0, True)
Else '内心
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
ang(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 1), _
                          C_display_wenti.m_point_no(num, 2), _
                          C_display_wenti.m_point_no(num, 0), 0, 0))
 ang(1) = Abs(angle_number(C_display_wenti.m_point_no(num, 3), _
                           C_display_wenti.m_point_no(num, 2), _
                           C_display_wenti.m_point_no(num, 0), 0, 0))
  If ang(0) > 0 And ang(1) > 0 Then
    tn_% = 0
    Call set_three_angle_value(ang(0), ang(1), 0, "1", "-1", "0", "0", _
            0, temp_record, tn_%, 0, 0, 0, 0, 0, True)
     Call C_display_wenti.set_m_condition_data(num, angle3_value_, tn_%)
  End If
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
   ang(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 2), _
                             C_display_wenti.m_point_no(num, 3), _
                             C_display_wenti.m_point_no(num, 0), 0, 0))
   ang(1) = Abs(angle_number(C_display_wenti.m_point_no(num, 1), _
                             C_display_wenti.m_point_no(num, 3), _
                             C_display_wenti.m_point_no(num, 0), 0, 0))
If ang(0) > 0 And ang(1) > 0 Then
            Call set_three_angle_value(ang(0), ang(1), 0, "1", "-1", "0", "0", _
             0, temp_record, 0, 0, 0, 0, 0, 0, True)
End If
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
   ang(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 3), _
                             C_display_wenti.m_point_no(num, 1), _
                             C_display_wenti.m_point_no(num, 0), 0, 0))
   ang(1) = Abs(angle_number(C_display_wenti.m_point_no(num, 2), _
                             C_display_wenti.m_point_no(num, 1), _
                             C_display_wenti.m_point_no(num, 0), 0, 0))
If ang(0) > 0 And ang(1) > 0 Then
            Call set_three_angle_value(ang(0), ang(1), 0, "1", "-1", "0", "0", _
            0, temp_record, 0, 0, 0, 0, 0, 0, True)
End If
Call C_display_wenti.set_m_point_no(num, _
      m_circle_number(1, 0, pointapi0, C_display_wenti.m_point_no(num, 4), _
        C_display_wenti.m_point_no(num, 5), _
          C_display_wenti.m_point_no(num, 6), 0, _
           0, 0, 1, 1, condition, condition_color, True), 10, False)
If C_display_wenti.m_point_no(num, 6) > 0 Then
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
k% = line_number(C_display_wenti.m_point_no(num, 2), _
                 C_display_wenti.m_point_no(num, 3), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                 condition, condition_color, 1, 0)
Call set_tangent_line(k%, _
       C_display_wenti.m_point_no(num, 4), _
          C_display_wenti.m_point_no(num, 10), 0, 0, temp_record, 0, 0)
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
tl(2) = line_number(C_display_wenti.m_point_no(num, 0), _
                    C_display_wenti.m_point_no(num, 4), _
                    pointapi0, pointapi0, _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 4)), _
                    condition, condition_color, 1, 0)
Call set_dverti(k%, tl(2), temp_record, 0, 0, True)
Call set_element_depend(point_, C_display_wenti.m_point_no(num, 4), _
                                   line_, k%, line_, tl(2), 0, 0, False)
'***************
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
k% = line_number(C_display_wenti.m_point_no(num, 1), _
                 C_display_wenti.m_point_no(num, 2), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                 condition, condition_color, 1, 0)
Call set_tangent_line(k%, _
       C_display_wenti.m_point_no(num, 6), C_display_wenti.m_point_no(num, 10), 0, 0, temp_record, 0, 0)
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
tl(2) = line_number(C_display_wenti.m_point_no(num, 0), _
                    C_display_wenti.m_point_no(num, 6), _
                    pointapi0, pointapi0, _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 6)), _
                    condition, condition_color, 1, 0)
Call set_dverti(k%, tl(2), temp_record, 0, 0, True)
Call set_element_depend(point_, C_display_wenti.m_point_no(num, 6), _
                                   line_, k%, line_, tl(2), 0, 0, False)
'*******************
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
k% = line_number(C_display_wenti.m_point_no(num, 3), _
                 C_display_wenti.m_point_no(num, 1), _
                 pointapi0, pointapi0, _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 3)), _
                 depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                 condition, condition_color, 1, 0)
Call set_tangent_line(k%, _
       C_display_wenti.m_point_no(num, 5), C_display_wenti.m_point_no(num, 10), 0, 0, temp_record, 0, 0)
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
tl(2) = line_number(C_display_wenti.m_point_no(num, 0), _
                    C_display_wenti.m_point_no(num, 5), _
                    pointapi0, pointapi0, _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 0)), _
                    depend_condition(point_, C_display_wenti.m_point_no(num, 5)), _
                    condition, condition_color, 1, 0)
Call set_dverti(k%, tl(2), temp_record, 0, 0, True)
Call set_element_depend(point_, C_display_wenti.m_point_no(num, 5), _
                                   line_, k%, line_, tl(2), 0, 0, False)
'Call set_element_depend(circle_, C_display_wenti.m_point_no(num,num,10), _
          point_, C_display_wenti.m_point_no(num,num,0), _
           point_, C_display_wenti.m_point_no(num,num,1), _
            point_, C_display_wenti.m_point_no(num,num,2))
Call set_equal_dline(C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 6), _
      C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 5), 0, 0, 0, 0, 0, 0, 0, _
       temp_record, 0, 0, 0, 0, 0, False)
Call set_equal_dline(C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 6), _
      C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 4), 0, 0, 0, 0, 0, 0, 0, _
       temp_record, 0, 0, 0, 0, 0, False)
Call set_equal_dline(C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 5), _
      C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 4), 0, 0, 0, 0, 0, 0, 0, _
       temp_record, 0, 0, 0, 0, 0, False)
Call set_equal_dline(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 6), _
      C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 5), 0, 0, 0, 0, 0, 0, 0, _
       temp_record, 0, 0, 0, 0, 0, False)
Call set_equal_dline(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 6), _
      C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 4), 0, 0, 0, 0, 0, 0, 0, _
       temp_record, 0, 0, 0, 0, 0, False)
Call set_equal_dline(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 4), _
      C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 5), 0, 0, 0, 0, 0, 0, 0, _
       temp_record, 0, 0, 0, 0, 0, False)
       
 Else
   Call C_display_wenti.set_m_point_no(num, 0, 4, False)
 End If
End If
Case 22
'!_~~、!_~~是方程!_~x[up\\2]+!_~x+!_~的两个根
Call draw_picture22(num)
Case 23
'23 □、□、□、□四点共圆
Call draw_picture23(num)
Case 24
'24 □、□、□三点共线
Call draw_picture24(num)
Case 25, 27, 28
'25 □□＝□□
'27 □□∥□□
'28 □□⊥□□
Call draw_picture25_27_28(num)
'"线段□□和□□长相等，即｜□□｜＝｜□□｜"
'"□□平行于□□"
' "□□垂直于□□"
'i% = line_number(c_display_wenti.m_point_no(num,num,0), _
 'c_display_wenti.m_point_no(num,num,1), concl, display)
'Call add_point_to_con_line(c_display_wenti.m_point_no(num,num,2), i%)
Case 29
'29 点□位于线段□□的垂直平分线上
Call draw_picture29(num)
Case 26
'26点□是线段□□的中点
'点□是线段□□的中点
Call draw_picture26(num)
Case 31
'31□□/□□=!_~
Call draw_picture31(num)
Case 30
'∠□□□=∠□□□
Call draw_picture30(num)
conclusion_data(last_conclusion).wenti_no = num
last_conclusion = last_conclusion + 1
'operate_step(C_display_wenti.m_last_input_wenti_no).last_con_circle = last_conditions.last_cond(1).aid_circle_no
operate_step(C_display_wenti.m_last_input_wenti_no).last_con_line = last_conditions.last_cond(1).con_line_no
operate_step(C_display_wenti.m_last_input_wenti_no).last_conclusion = last_conclusion
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
'MDIForm1.add_point.Enabled = True
MDIForm1.Toolbar1.Buttons(19).visible = True
Case 32
'32 □□/□□＝□□/□□
Call draw_picture32(num)
Case 33, 34
'33 △□□□∽△□□□
'34△□□□≌△□□□
Call draw_picture33_34(num)
Case 35, 54
'35 □□=?
'54 □□=!_~
Call draw_picture35_54(num)
If C_display_wenti.m_no(num) = 35 Then
 If th_chose(-5).chose < 2 Then
 th_chose(-5).chose = 1
 End If
End If
Case 36, 53 '
'36 ∠□□□=?
'53 ∠□□□=!_~°
If C_display_wenti.m_no(num) = 36 Then
 If th_chose(-5).chose < 2 Then
 th_chose(-5).chose = 1
 End If
End If
Call draw_picture36_53(num)
Case 37
'□□/□□=?
If th_chose(-5).chose < 2 Then
th_chose(-5).chose = 1
End If
Call draw_picture37(num)
using_area_th = 8
Case 38, -49, 50, 74, 73, 75, 76
'-49  _~
'50_~=?
'38 _~
If C_display_wenti.m_no(num) = 50 Then
If th_chose(-5).chose < 2 Then
th_chose(-5).chose = 1
End If
End If
   Call draw_picture38_49(num, 0, 49, no_reduce)
Case 39
'□□、□□、□□三直线共点
Call draw_picture39(num)
Case 40
'40 △□□□是等边三角形
Call draw_picture40(num)
Case 41
'41 △□□□是等腰三角形
Call draw_picture41(num)
Case 42
'41 △□□□是等腰直角三角形
Call draw_picture42(num)
Case 43
' 43 □□□□是长方形
Call draw_picture43(num)
Case 44
'44 □□□□是正方形
Call draw_picture44(num)
Case 45
'45 □□□□是平行四边形
Call draw_picture45(num)
Case 46
'46 □□□□是菱形
Call draw_picture46(num)
Case 47
'47 直线□□与⊙□[down\\(_)]相切于□
Call draw_picture47(num)
Case 48
'□□□□是梯形
Case 49
'49 □□□□是等腰梯形
Call draw_picture49(num)
Case 51
'∠□□□=∠□□□+∠□□□
Call draw_picture51(num)
Case 52
'52 ∠□□□+∠□□□=!_~ °
Call draw_picture52(num)
Case 55
'55 ∠□□□/∠□□□=!_~
Call draw_picture55(num)
Case 56
'56 求△□□□的面积
If th_chose(-5).chose < 2 Then
th_chose(-5).chose = 1
End If
Call draw_picture56(num)
Case 57
'57 求四边形□□□□的面积
If th_chose(-5).chose < 2 Then
th_chose(-5).chose = 1
End If
Call draw_picture57(num)
Case 58
'58 求⊙□[down\\(_)]的面积
If th_chose(-5).chose < 2 Then
th_chose(-5).chose = 1
End If
Call draw_picture58(num)
Case 59
'59 扇形□□□的面积=?
If th_chose(-5).chose < 2 Then
th_chose(-5).chose = 1
End If
Call draw_picture59(num)
Case 60, 62
'60 △□□□的周长=?
'62 △□□□的周长=!_~
If th_chose(-5).chose < 2 Then
th_chose(-5).chose = 1
End If
Call draw_picture60_62(num)
Case 61
'61 求⊙□[down\\(_)]的周长
If th_chose(-5).chose < 2 Then
th_chose(-5).chose = 1
Call draw_picture61(num)
End If
'Case 62
'Call draw_picture60_62(num)
Case 63
'63 △□□□的面积=!_~
Call draw_picture63(num)
Case 64, 66
'64 四边形□□□□的周长=!_~
'66 四边形□□□□的周长=?
If C_display_wenti.m_no(num) = 66 Then
If th_chose(-5).chose < 2 Then
th_chose(-5).chose = 1
End If
End If
Call draw_picture64_66(num)
Case 65
'65 四边形□□□□的面积=!_~
Call draw_picture65(num)
'67 求以!_~和!_~为根的一元二次方程
Case 68, 69
'68 △□□□与△□□□面积相等
'69 ∠□□□/∠□□□=!_~
If C_display_wenti.m_no(num) = 69 Then
If th_chose(-5).chose < 2 Then
th_chose(-5).chose = 1
End If
End If
Call draw_picture68_69(num)
'70 直线□□与⊙□□□相切于□
'71⊙□□□的面积=?
'72 ⊙□□□的周长?
'73 _~=定值
End Select
'Call set_data_inform(num)
MDIForm1.Timer1.Enabled = False
draw_wenti_no = num + 1
'If event_statue <> wait_for_input_char Then
' event_statue = ready
'End If
'If c_display_wenti.m_no(num) < 23 Then
'Call call_theorem(0)
'End If
End Sub
Private Sub set_initial_condition_56(ByVal num As Integer, temp_record As total_record_type)
Dim i%, tn_%, tp%
Dim it(3) As Integer
Dim c_data As condition_data_type
Dim value As String
Call set_item0(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
   0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, _
           it(0), 0, 0, c_data, False)
Call set_item0(C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
   0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, _
           it(1), 0, 0, c_data, False)
Call set_item0(C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 1), _
   0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, _
           it(2), 0, 0, c_data, False)
Call set_item0(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 3), _
   0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, _
           it(3), 0, 0, c_data, False)
    value = initial_string(number_string(C_display_wenti.m_point_no(num, 4))) 'initial_string(cond_to_string(num%, 4, 18, 0))
      If InStr(1, value, ".", 0) > 0 Then
        th_chose(-5).chose = 2
      End If
'Call read_number_from_wenti(num%, 4, 0, 0, v$)
tn_% = 0
Call set_general_string(it(0), it(1), it(2), it(3), "1", "1", _
       "1", "1", value, 0, 0, 1, temp_record, tn_%, 0)
      Call C_display_wenti.set_m_condition_data(num, general_string_, tn_%)
tp% = C_display_wenti.m_point_no(num, 0)
For i% = 1 To 3
 If tp% < C_display_wenti.m_point_no(num, i%) Then
     tp% = C_display_wenti.m_point_no(num, i%)
 End If
Next i%

End Sub

Private Sub set_initial_condition_54_53_25(ByVal num As Integer, temp_record As total_record_type)
'-54 □□的垂直平分线交□□于□
'-53 □□的垂直平分线交⊙□[down\\(_)]于□
'-25 □□的垂直平分线交⊙□□□于□
Dim tl(2) As Integer
Dim tn_%
Dim dir(1) As String
Dim it(1) As Integer
Dim para(1) As String
Dim c_data As condition_data_type
'depend_no(num) = 1
 tn_% = 0
   tl(0) = C_display_wenti.m_inner_lin(num, 1) '相交线
   tl(1) = C_display_wenti.m_inner_lin(num, 2) '垂直平分线
   tl(2) = C_display_wenti.m_inner_lin(num, 3) '□□
Call set_mid_point(C_display_wenti.m_point_no(num, 0), _
                      C_display_wenti.m_inner_poi(num, 2), _
                       C_display_wenti.m_point_no(num, 1), 0, 0, 0, 0, 0, _
                        temp_record, 0, 0, 0, 0, 0)
 Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 2), _
          point_, C_display_wenti.m_point_no(num, 0), _
           point_, C_display_wenti.m_point_no(num, 1), 0, 0, True)  '中点
 Call set_element_depend(line_, tl(1), _
          point_, C_display_wenti.m_inner_poi(num, 2), _
           line_, tl(2), 0, 0, True) '中垂线
    Call set_dverti(tl(0), tl(1), temp_record, 0, 0, True)
    Call set_verti_mid_line(C_display_wenti.m_point_no(num, 0), _
                             C_display_wenti.m_inner_poi(num, 2), _
                               C_display_wenti.m_point_no(num, 1), _
                                  C_display_wenti.m_inner_lin(num, 1), temp_record, 0, 0)
   If C_display_wenti.m_no(num) = -54 Then
    Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
                             line_, C_display_wenti.m_inner_lin(num, 1), _
                              line_, C_display_wenti.m_inner_lin(num, 2), 0, 0, False)
   Else
    Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
                             line_, C_display_wenti.m_inner_lin(num, 2), _
                              circle_, C_display_wenti.m_inner_circ(num, 1), 0, 0, False)
   End If
If regist_data.run_type = 1 Then
   Call set_V_coordinate_system(C_display_wenti.m_point_no(num, 0), _
              C_display_wenti.m_point_no(num, 1))
   Call set_V_coordinate_system(C_display_wenti.m_point_no(num, 2), _
              C_display_wenti.m_point_no(num, 2))
   tl(0) = vector_number(C_display_wenti.m_point_no(num, 0), _
                 C_display_wenti.m_point_no(num, 4), dir(0))
   tl(1) = vector_number(C_display_wenti.m_point_no(num, 1), _
                 C_display_wenti.m_point_no(num, 4), dir(1))
   Call set_item0(tl(0), -10, tl(0), -10, "*", 0, 0, 0, 0, 0, "1", "1", "1", "1", "", para(0), 0, _
          c_data, 0, it(0), 0, 0, temp_record.record_data.data0.condition_data, False)
   Call set_item0(tl(1), -10, tl(1), -10, "*", 0, 0, 0, 0, 0, "1", "1", "1", "1", "", para(1), 0, _
          c_data, 0, it(1), 0, 0, temp_record.record_data.data0.condition_data, False)
   Call set_general_string(it(0), it(1), 0, 0, para(0), time_string("-1", para(1), True, False), _
            "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
End If
If regist_data.run_type = 1 Then
   Call set_V_coordinate_system(C_display_wenti.m_point_no(num, 0), _
                                    C_display_wenti.m_point_no(num, 1))
   Call set_V_coordinate_system( _
        m_Circ(C_display_wenti.m_point_no(num, 42)).data(0).data0.center, _
         m_Circ(C_display_wenti.m_point_no(num, 42)).data(0).data0.in_point(1))
   Call add_new_point_on_circle_for_vector(C_display_wenti.m_point_no(num, 42), _
              C_display_wenti.m_point_no(num, 44), temp_record)
    tl(0) = vector_number(C_display_wenti.m_point_no(num, 0), _
                          C_display_wenti.m_point_no(num, 44), dir(0))
    tl(1) = vector_number(C_display_wenti.m_point_no(num, 1), _
                          C_display_wenti.m_point_no(num, 44), dir(1))
    Call set_item0(tl(0), -10, tl(0), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", _
           para(0), 0, c_data, 0, it(0), 0, 0, temp_record.record_data.data0.condition_data, False)
    Call set_item0(tl(1), -10, tl(1), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", _
           para(0), 0, c_data, 0, it(1), 0, 0, temp_record.record_data.data0.condition_data, False)
    Call set_general_string(tl(0), tl(1), 0, 0, para(0), time_string("-1", para(1), True, False), _
           "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
End If
End Sub

Private Sub set_initial_condition_52_51_56(ByVal num As Integer, temp_record As total_record_type)
'-51∠□□□的平分线交□□于□
'-52∠□□□的平分线交⊙□[down\\(_)]于□
'-56∠□□□的平分线交⊙□□□于□
Dim triA(1) As Integer
Dim tl(2) As Integer
Dim tn_%
Dim tp%
'depend_no(num) = 1
'If C_display_wenti.m_no(num) = -56 Then
     tp% = C_display_wenti.m_inner_poi(num, 1)
     tl(0) = C_display_wenti.m_inner_lin(num, 3)
     tl(1) = C_display_wenti.m_inner_lin(num, 4)
     tl(2) = C_display_wenti.m_inner_lin(num, 2)
     Call set_element_depend(line_, tl(2), line_, tl(0), line_, tl(1), 0, 0, False)
     '角平分线由边确定
    triA(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 0), _
         C_display_wenti.m_point_no(num, 1), tp%, 0, 0))
    triA(1) = Abs(angle_number(tp%, C_display_wenti.m_point_no(num, 1), _
              C_display_wenti.m_point_no(num, 2), 0, 0))
    If triA(0) > 0 And triA(1) > 0 Then '平分线
     tn_% = 0
      Call set_three_angle_value(triA(0), triA(1), 0, "1", "-1", "0", "0", _
               0, temp_record, tn_%, 0, 0, 0, 0, 0, True)
               Call C_display_wenti.set_m_condition_data(num, angle3_value_, tn_%)
    End If
 If C_display_wenti.m_no(num) = -51 Then
    tl(0) = C_display_wenti.m_inner_lin(num, 1)
         Call set_element_depend(point_, tp%, line_, tl(2), line_, tl(0), 0, 0, False)
          '交点
 Else
             Call set_element_depend(point_, tp%, line_, tl(2), circle_, _
                         C_display_wenti.m_inner_circ(num, 1), 0, 0, False)
 End If
 If regist_data.run_type = 1 Then
    Call add_new_point_on_circle_for_vector(C_display_wenti.m_point_no(num, 42), _
                 C_display_wenti.m_point_no(num, 44), temp_record)
 End If
End Sub

Private Sub set_initial_condition_50(ByVal num As Integer, temp_record As total_record_type)
'-50□□是∠□□□的平分线
Dim triA(1) As Integer
Dim tl(1) As Integer
Dim tp(4) As Integer
Dim tn_%, i%
If C_display_wenti.m_inner_poi(num, 1) = 0 Then
tp(0) = C_display_wenti.m_point_no(num, 0) '∠012
tp(1) = C_display_wenti.m_point_no(num, 1)
tp(2) = C_display_wenti.m_point_no(num, 2)
tp(3) = C_display_wenti.m_point_no(num, 3)
tp(4) = C_display_wenti.m_point_no(num, 4)
If tp(0) = tp(3) Then
   tp(0) = tp(1)
End If
 Call set_wenti_cond_50(tp(0), tp(1), tp(2), tp(3), num)
End If
    triA(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 0), _
         C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 3), 0, 0))
    triA(1) = Abs(angle_number(C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 1), _
            C_display_wenti.m_point_no(num, 3), 0, 0))
If triA(0) > 0 And triA(1) > 0 Then
 tn_% = 0
       Call set_three_angle_value(triA(0), triA(1), 0, "1", "-1", "0", "0", _
              0, temp_record, tn_%, 0, 0, 0, 0, 0, True)
              Call C_display_wenti.set_m_condition_data(num, angle3_value_, tn_%)
End If
tl(0) = C_display_wenti.m_inner_lin(num, 3)
tl(1) = C_display_wenti.m_inner_lin(num, 4)
If C_display_wenti.m_no_(num) = -50 Then
 Call set_element_depend(line_, C_display_wenti.m_inner_lin(num, 2), _
                          point_, C_display_wenti.m_inner_poi(num, 3), _
                           line_, C_display_wenti.m_inner_lin(num, 3), _
                            line_, C_display_wenti.m_inner_lin(num, 4), False)
 If C_display_wenti.m_inner_lin(num, 1) > 0 Then
  Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
                          line_, C_display_wenti.m_inner_lin(num, 2), _
                           line_, C_display_wenti.m_inner_lin(num, 1), 0, 0, True)
 ElseIf C_display_wenti.m_inner_circ(num, 1) > 0 Then
  Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
                         line_, C_display_wenti.m_inner_lin(num, 2), _
                          circle_, C_display_wenti.m_inner_circ(num, 1), 0, 0, True)
 Else
  Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
                                   line_, C_display_wenti.m_inner_lin(num, 2), _
                                    0, 0, 0, 0, False)
 End If
ElseIf C_display_wenti.m_no_(num) = -501 Then
 Call set_element_depend(line_, C_display_wenti.m_inner_lin(num, 3), _
                         line_, C_display_wenti.m_inner_lin(num, 2), _
                          line_, C_display_wenti.m_inner_lin(num, 4), _
                           0, 0, False)
 If C_display_wenti.m_inner_lin(num, 1) > 0 Then
  Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 2), _
                         line_, C_display_wenti.m_inner_lin(num, 3), _
                          line_, C_display_wenti.m_inner_lin(num, 1), 0, 0, True)
 ElseIf C_display_wenti.m_inner_circ(num, 1) > 0 Then
   Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 2), _
                         line_, C_display_wenti.m_inner_lin(num, 3), _
                          circle_, C_display_wenti.m_inner_circ(num, 1), 0, 0, True)
Else
 Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 2), _
                         line_, C_display_wenti.m_inner_lin(num, 3), _
                          0, 0, 0, 0, False)
 End If
ElseIf C_display_wenti.m_no_(num) = -502 Then 'If C_display_wenti.m_inner_point_type <= 2 Then
 Call set_element_depend(line_, C_display_wenti.m_inner_lin(num, 4), _
                          line_, C_display_wenti.m_inner_lin(num, 3), _
                           line_, C_display_wenti.m_inner_lin(num, 2), _
                            0, 0, False)
 If C_display_wenti.m_inner_lin(num, 1) > 0 Then
  Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 4), _
                           line_, C_display_wenti.m_inner_lin(num, 4), _
                            line_, C_display_wenti.m_inner_lin(num, 1), 0, 0, True)
 ElseIf C_display_wenti.m_inner_circ(num, 1) > 0 Then
  Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 4), _
                           line_, C_display_wenti.m_inner_lin(num, 4), _
                            circle_, C_display_wenti.m_inner_circ(num, 1), 0, 0, True)
 Else
  Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 4), _
                           line_, C_display_wenti.m_inner_lin(num, 4), _
                            0, 0, 0, 0, False)
 End If
End If
End Sub

Private Sub set_initial_condition_48_47(ByVal num As Integer, temp_record As total_record_type)
Dim value As String
Dim triA As Integer
Dim tn_%
Dim tl(1) As Integer
Dim dir(1) As String
Dim c_data As condition_data_type
value = initial_string(number_string(C_display_wenti.m_point_no(num, 3)))
triA = triangle_number(C_display_wenti.m_point_no(num, 0), _
    C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
       0, 0, 0, 0, 0, 0, 0)
If C_display_wenti.m_no(num) = -47 Then
tn_% = 0
Call set_area_of_triangle(triA, value, temp_record, tn_%, 0)
            Call C_display_wenti.set_m_condition_data(num, area_of_element_, tn_%)
If InStr(1, value, ".", 0) > 0 Then
 th_chose(-5).chose = 2
End If
area_of_triangle_conclusion = 1
Else
tn_% = 0
Call set_three_line_value(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 1), _
   C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 2), _
    C_display_wenti.m_point_no(num, 0), 0, 0, 0, 0, 0, 0, 0, 0, 0, _
     "1", "1", "1", value, temp_record, tn_%, 0, 0)
     Call C_display_wenti.set_m_condition_data(num, line3_value_, tn_%)
End If
If regist_data.run_type = 1 Then
   If C_display_wenti.m_no(num) = -47 Then
      Call set_V_coordinate_system(C_display_wenti.m_point_no(num, 0), _
                                      C_display_wenti.m_point_no(num, 1))
      Call set_V_coordinate_system(C_display_wenti.m_point_no(num, 0), _
                                      C_display_wenti.m_point_no(num, 2))
      tl(0) = vector_number(C_display_wenti.m_point_no(num, 0), _
                           C_display_wenti.m_point_no(num, 1), dir(0))
      tl(1) = vector_number(C_display_wenti.m_point_no(num, 0), _
                           C_display_wenti.m_point_no(num, 2), dir(1))
      temp_record.record_data.data0.condition_data.condition_no = 0
      Call add_conditions_to_record(area_of_element_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
      Call set_item0(tl(0), -10, tl(1), -10, "*", 0, 0, 0, 0, 0, 0, _
                      "1", "1", "1", time_string(dir(1), time_string(dir(0), _
                         time_string("2", area_of_element(tn_%).data(0).value, False, False), False, False), True, False), _
                          "", 0, temp_record.record_data.data0.condition_data, 0, 0, 0, 0, c_data, False)
   Else
   End If
End If
End Sub

Private Sub set_initial_condition_46_45(ByVal num As Integer, temp_record As total_record_type)
Dim value As String
Dim it(3) As Integer
Dim tn_%
Dim c_data As condition_data_type
 value = initial_string(number_string(C_display_wenti.m_point_no(num, 4)))
If C_display_wenti.m_no(num) = -45 Then
tn_% = 0
Call set_area_of_polygon(C_display_wenti.m_point_no(num, 0), _
   C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
    C_display_wenti.m_point_no(num, 3), value, temp_record, tn_%, 0)
    Call C_display_wenti.set_m_condition_data(num, area_of_element_, tn_%)
    area_of_triangle_conclusion = 1
     using_area_th = 8
      If InStr(1, value, ".", 0) > 0 Then
        th_chose(-5).chose = 2
      End If
Else
Call set_item0(C_display_wenti.m_point_no(num, 0), _
     C_display_wenti.m_point_no(num, 1), 0, 0, _
      "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, _
        it(0), 0, 0, c_data, False)
Call set_item0(C_display_wenti.m_point_no(num, 1), _
     C_display_wenti.m_point_no(num, 2), 0, 0, _
      "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, _
        it(1), 0, 0, c_data, False)
Call set_item0(C_display_wenti.m_point_no(num, 2), _
     C_display_wenti.m_point_no(num, 3), 0, 0, _
      "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, _
        it(2), 0, 0, c_data, False)
Call set_item0(C_display_wenti.m_point_no(num, 3), _
     C_display_wenti.m_point_no(num, 0), 0, 0, _
      "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, _
        it(3), 0, 0, c_data, False)
tn_% = 0
Call set_general_string(it(0), it(1), it(2), it(3), _
   "1", "1", "1", "1", value, 0, 0, 1, temp_record, tn_%, 0)
      If InStr(1, value, ".", 0) > 0 Then
      Call C_display_wenti.set_m_condition_data(num, general_string_, tn_%)
        th_chose(-5).chose = 2
      End If
End If
End Sub
Private Sub set_initial_condition4(ByVal num As Integer, temp_record As total_record_type)
'4 在□□的垂直平分线上任取一点□
Dim tn_%, no_%
Dim tl(1) As Integer
'depend_no(num) = 1
tn_% = 0
If C_display_wenti.m_inner_poi(num, 1) = 0 Then
   Call set_wenti_cond4( _
           C_display_wenti.m_point_no(num, 0), _
           C_display_wenti.m_point_no(num, 1), _
           C_display_wenti.m_point_no(num, 2), _
            0, 0, num)
End If
  tl(0) = C_display_wenti.m_inner_lin(num, 2)
  tl(1) = C_display_wenti.m_inner_lin(num, 1)
  Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 2), _
         point_, C_display_wenti.m_point_no(num, 0), _
          point_, C_display_wenti.m_point_no(num, 1), _
           0, 0, True) '中点
  Call set_element_depend(line_, tl(0), _
            point_, C_display_wenti.m_inner_poi(num, 2), _
             line_, C_display_wenti.m_inner_lin(num, 3), _
                 0, 0, False) '垂直平分线
  no_% = C_display_wenti.m_no_(num)
  If no_% = 0 Then
  Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
            line_, tl(1), 0, 0, 0, 0, False) '垂直平分线上的点
  ElseIf no_% = -54 Then
  Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
            line_, tl(0), line_, tl(1), 0, 0, True) '垂直平分线上的点
  ElseIf no_% = -53 Then
  Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
            line_, tl(0), circle_, C_display_wenti.m_inner_circ(num, 1), 0, 0, True) '垂直平分线上的点
  End If
  Call set_verti_mid_line(C_display_wenti.m_point_no(num, 0), _
                          C_display_wenti.m_inner_poi(num, 2), _
                           C_display_wenti.m_point_no(num, 1), _
                            tl(1), temp_record, tn_%, 0)
  Call C_display_wenti.set_m_condition_data(num, verti_mid_line_, tn_%)
End Sub

Private Sub set_initial_condition_43_42(ByVal num As Integer, temp_record As total_record_type)
'-42 在⊙□[down\\(_)]上取一点□使得□□＝!_~
'-43 在□□上取一点□使得□□＝_"
'-57 在⊙□□□上取一点□使得□□＝!_~
Dim value As String
Dim tn_%
Dim tl(1) As Integer
If C_display_wenti.m_no(num) <> -57 Then
   value = initial_string(number_string(C_display_wenti.m_point_no(num, 4)))
Else
   value = initial_string(number_string(C_display_wenti.m_point_no(num, 6)))
End If
tn_% = 0
  Call set_line_value(C_display_wenti.m_inner_poi(num, 2), _
     C_display_wenti.m_inner_poi(num, 3), _
    value, 0, 0, 0, temp_record.record_data, tn_%, 0, True)
    Call C_display_wenti.set_m_condition_data(num, line_value_, tn_%)
      If InStr(1, value, ".", 0) > 0 Then
        th_chose(-5).chose = 2
      End If
 '  Call set_element_depend(circle_, C_display_wenti.m_inner_circ(num, 2), _
                     point_, _
                     m_Circ(C_display_wenti.m_inner_circ(num, 2)).data(0).data0.center, _
                     0, 0, 0, 0, False) '定半径的圆
 '  m_Circ(C_display_wenti.m_inner_circ(num, 2)).data(0).radii_depend_poi(0) = -1
If C_display_wenti.m_no(num) = -43 Then
  ' Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
       line_, C_display_wenti.m_inner_lin(num, 1), circle_, _
        C_display_wenti.m_inner_circ(num, 2), 0, 0, False)
Else
'   Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
       circle_, C_display_wenti.m_inner_circ(num, 1), _
        circle_, C_display_wenti.m_inner_circ(num, 2), 0, 0, False)
End If
End Sub

Private Sub set_initial_condition_41(ByVal num As Integer, temp_record As total_record_type)
Dim ang(1) As Integer
Dim value As String
Dim tn_%
ang(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 0), _
  C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), 0, 0))
ang(1) = Abs(angle_number(C_display_wenti.m_point_no(num, 3), _
  C_display_wenti.m_point_no(num, 4), C_display_wenti.m_point_no(num, 5), 0, 0))
value = initial_string(number_string(C_display_wenti.m_point_no(num, 6))) 'initial_string(cond_to_string(num, 6, 18, 0))
      If InStr(1, value, ".", 0) > 0 Then
        th_chose(-5).chose = 2
      End If
If ang(0) > 0 And ang(1) > 0 Then
  tn_% = 0
   Call set_angle_relation(ang(0), ang(1), value, "1", _
      temp_record, tn_%, 0, True)
      Call C_display_wenti.set_m_condition_data(num, angle_relation_, tn_%)
End If

End Sub

Private Sub set_initial_condition_40(ByVal num As Integer, temp_record As total_record_type)
Dim it(3) As Integer
Dim tn(7) As Integer
Dim tn_%
it(0) = line_number0(C_display_wenti.m_point_no(num, 0), _
       C_display_wenti.m_point_no(num, 1), tn(0), tn(1))
it(1) = line_number0(C_display_wenti.m_point_no(num, 2), _
       C_display_wenti.m_point_no(num, 3), tn(2), tn(3))
it(2) = line_number0(C_display_wenti.m_point_no(num, 4), _
       C_display_wenti.m_point_no(num, 5), tn(4), tn(5))
it(3) = line_number0(C_display_wenti.m_point_no(num, 6), _
       C_display_wenti.m_point_no(num, 7), tn(6), tn(7))
If it(0) > 0 And it(1) > 0 And it(2) > 0 And it(3) > 0 Then
tn_% = 0
Call set_dpoint_pair(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 4), _
   C_display_wenti.m_point_no(num, 5), C_display_wenti.m_point_no(num, 6), _
    C_display_wenti.m_point_no(num, 7), tn(0), tn(1), tn(2), tn(3), _
     tn(4), tn(5), tn(6), tn(7), it(0), it(1), it(2), it(3), _
      0, temp_record, True, tn_%, 0, 0, 0, True)
      Call C_display_wenti.set_m_condition_data(num, dpoint_pair_, tn_%)
Else
  error_of_wenti = 4
End If

End Sub

Private Sub set_initial_condition_39(ByVal num As Integer, temp_record As total_record_type)
Dim ang(2) As Integer
Dim tn_%
'inpcond(-39) = ∠□□□=∠□□□+∠□□□
ang(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), 0, 0))
 ang(1) = Abs(angle_number(C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 4), _
     C_display_wenti.m_point_no(num, 5), 0, 0))
  ang(2) = Abs(angle_number(C_display_wenti.m_point_no(num, 6), _
      C_display_wenti.m_point_no(num, 7), C_display_wenti.m_point_no(num, 8), 0, 0))
 If ang(0) > 0 And ang(1) > 0 And ang(2) > 0 Then
  tn_% = 0
  Call set_three_angle_value(ang(0), ang(1), ang(2), _
        "1", "-1", "-1", "0", 0, temp_record, tn_%, 0, 0, 0, 0, 0, True)
        Call C_display_wenti.set_m_condition_data(num, angle3_value_, tn_%)
End If

End Sub

Private Sub set_initial_condition_38(ByVal num As Integer, temp_record As total_record_type)
Dim value As String
Dim ang(1) As Integer
Dim tn_%
value = initial_string(number_string(C_display_wenti.m_point_no(num, 6))) 'initial_string(cond_to_string(num, 6, 18, 0))
      If InStr(1, value, ".", 0) > 0 Then
        th_chose(-5).chose = 2
      End If
ang(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), 0, 0))
ang(1) = Abs(angle_number(C_display_wenti.m_point_no(num, 3), _
 C_display_wenti.m_point_no(num, 4), C_display_wenti.m_point_no(num, 5), 0, 0))
  If ang(0) <> 0 And ang(1) <> 0 Then
   tn_% = 0
    Call set_three_angle_value(Abs(ang(0)), Abs(ang(1)), 0, "1", "1", "0", value, _
          0, temp_record, tn_%, 0, 0, 0, 0, 0, True)
          Call C_display_wenti.set_m_condition_data(num, angle3_value_, tn_%)
  End If

End Sub

Private Sub set_initial_condition_37(ByVal num As Integer, temp_record As total_record_type)
Dim tn_%
temp_record.record_data.data0.condition_data.condition_no = 0 ' = record0
tn_% = 0
Call set_total_equal_triangle(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 4), _
   C_display_wenti.m_point_no(num, 5), temp_record, tn_%, 0)
   Call C_display_wenti.set_m_condition_data(num, total_equal_triangle_, tn_%)

End Sub

Private Sub set_initial_condition_36(ByVal num As Integer, temp_record As total_record_type)
Dim tn_%
temp_record.record_data.data0.condition_data.condition_no = 0 ' = record0
tn_% = 0
Call set_similar_triangle(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 4), _
   C_display_wenti.m_point_no(num, 5), temp_record, tn_%, 0, 0)
    Call C_display_wenti.set_m_condition_data(num, similar_triangle_, tn_%)

End Sub

Private Sub set_initial_condition_35(ByVal num As Integer, temp_record As total_record_type)
Dim tn_%
temp_record.record_data.data0.condition_data.condition_no = 0 ' = record0
tn_% = 0
Call set_three_line_value(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), C_display_wenti.m_point_no(num, 4), _
   C_display_wenti.m_point_no(num, 5), 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    "1", "-1", "-1", "0", temp_record, tn_%, 0, 0)
    Call C_display_wenti.set_m_condition_data(num, line3_value_, tn_%)

End Sub

Private Sub set_initial_condition_34(ByVal num As Integer, temp_record As total_record_type)
Dim value As String
Dim tn_%
value = initial_string(number_string(C_display_wenti.m_point_no(num, 4))) 'initial_string(cond_to_string(num, 4, 18, 0))
      If InStr(1, value, ".", 0) > 0 Then
        th_chose(-5).chose = 2
      End If
temp_record.record_data.data0.condition_data.condition_no = 0 ' = record0
tn_% = 0
Call set_two_line_value(C_display_wenti.m_point_no(num, 0), _
    C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
      C_display_wenti.m_point_no(num, 3), 0, 0, 0, 0, 0, 0, _
       "1", "1", value, temp_record, tn_%, 0)
       Call C_display_wenti.set_m_condition_data(num, two_line_value_, tn_%)

End Sub

Private Sub set_initial_condition_33_44(ByVal num As Integer, temp_record As total_record_type)
'-33过□作⊙□[down\\(_)]的切线□□
'-44过□作⊙□□□的切线□□
Dim i%, j%
Dim tn_%, tan_p%
Dim t_lv(2) As Integer
Dim tn(2) As Integer
Dim para(2) As String
Dim c_data As condition_data_type
' □□⊥□□
If C_display_wenti.m_inner_point_type(num) = 0 Then
   tan_p% = C_display_wenti.m_inner_poi(num, 2)
Else
   tan_p% = C_display_wenti.m_inner_poi(num, 1)
End If
t_line(1) = C_display_wenti.m_inner_lin(num, 1)
temp_record.record_data.data0.condition_data.condition_no = 0 ' = record0
tn_% = 0
Call set_tangent_line(t_line(1), tan_p%, _
             C_display_wenti.m_inner_circ(num, 1), _
              0, 0, temp_record, tn_%, 0)
   Call C_display_wenti.set_m_condition_data(num, tangent_line_, tn_%)
temp_record.record_data.data0.condition_data.condition_no = 0 ' = record0
If C_display_wenti.m_inner_point_type(num) = 0 Then
    Call set_element_depend(line_, t_line(1), point_, _
         C_display_wenti.m_inner_poi(num, 2), _
           circle_, C_display_wenti.m_inner_circ(num, 1), 0, 0, False)
    Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
               line_, C_display_wenti.m_inner_lin(num, 1), _
                    0, 0, 0, 0, False)
Else
    Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
               point_, C_display_wenti.m_inner_poi(num, 2), _
                circle_, C_display_wenti.m_inner_circ(num, 1), 0, 0, True)
    Call set_element_depend(line_, t_line(1), _
            point_, C_display_wenti.m_inner_poi(num, 2), _
              point_, C_display_wenti.m_inner_poi(num, 1), 0, 0, False)
End If
If regist_data.run_type = 1 Then
   Call set_V_coordinate_system(m_Circ(C_display_wenti.m_point_no(num, 42)).data(0).data0.center, _
          C_display_wenti.m_point_no(num, 0))
   t_lv(0) = vector_number(C_display_wenti.m_point_no(num, 35), _
                    C_display_wenti.m_point_no(num, 36), 0)
   t_lv(1) = vector_number(C_display_wenti.m_point_no(num, 35), _
                    m_Circ(C_display_wenti.m_point_no(num, 42)).data(0).data0.center, 0)
   t_lv(2) = vector_number(C_display_wenti.m_point_no(num, 36), _
                    m_Circ(C_display_wenti.m_point_no(num, 42)).data(0).data0.center, 0)
   Call set_item0(t_lv(0), -10, t_lv(1), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "0", "1", 0, _
          temp_record.record_data.data0.condition_data, 0, 0, 0, 0, c_data, False)
   Call set_item0(t_lv(0), -10, t_lv(0), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), 0, _
           c_data, 0, tn(0), 0, 0, c_data, False)
   Call set_item0(t_lv(1), -10, t_lv(1), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), 0, _
          c_data, 0, tn(1), 0, 0, c_data, False)
   Call set_item0(t_lv(2), -10, t_lv(2), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(2), 0, _
          c_data, 0, tn(2), 0, 0, c_data, False)
   temp_record.record_data.data0.condition_data.condition_no = 0
   'temp_record.record_.display_no = 0
   Call add_conditions_to_record(tangent_line_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
   tn_% = 0
   Call set_general_string(tn(0), tn(1), tn(2), 0, para(0), para(1), _
            time_string("-1", para(2), True, False), "0", "0", 0, _
              0, 0, temp_record, tn_%, 0)
   Call set_item0(t_lv(0), -10, t_lv(1), -10, "*", 0, 0, 0, 0, 0, 0, _
         "1", "1", "1", "0", para(0), 0, temp_record.record_data.data0.condition_data, _
           0, tn(0), 0, 0, c_data, False)
   tn_% = 0
   Call set_general_string(tn(0), 0, 0, 0, para(0), "0", "0", _
             "0", "0", 0, 0, 0, temp_record, tn_%, 0)
   temp_record.record_data.data0.condition_data.condition_no = 0
   Call add_conditions_to_record(tangent_line_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
   Call add_new_point_on_circle_for_vector(C_display_wenti.m_point_no(num, 42), _
              C_display_wenti.m_point_no(num, 35), temp_record)
End If
End Sub

Private Sub set_initial_condition_32_3(ByVal num As Integer, temp_record As total_record_type)
'-3 与⊙□[down\\(_)]相切于点□的切线交⊙□[down\\(_)]
'-32 与⊙□[down\\(_)]相切于点□的切线交直线□□于□
'-29 与⊙□□□相切于□的切线交⊙□[down\\(_)]于□
'-28 与⊙□□□相切于□的切线交⊙□□□于□
'-27 与⊙□□□相切于□的切线交直线□□于□
'-26 与⊙□[down\\(_)]相切于□的切线交⊙□□□于□
Dim tn_%, tp%
Dim tl(1) As Integer
Dim p_coord As POINTAPI
Dim wenti_no%
Dim c_data0 As condition_data_type
'inpcond(-32) = 与⊙□(_)相切于点□的切线交直线□□于□ '□□∥□□
temp_record.record_data.data0.condition_data.condition_no = 0 ' = record0
wenti_no% = C_display_wenti.m_no(num)
If wenti_no = -32 Or wenti_no = -27 Then
t_line(0) = C_display_wenti.m_inner_lin(num, 2)
tn_% = 0
Call set_tangent_line(t_line(0), C_display_wenti.m_inner_poi(num, 2), _
                        C_display_wenti.m_inner_circ(num, 1), _
                          0, 0, temp_record, tn_%, 0)
Call set_element_depend(line_, C_display_wenti.m_inner_lin(num, 2), _
               point_, C_display_wenti.m_inner_poi(num, 2), _
                circle_, C_display_wenti.m_inner_circ(num, 1), _
                 0, 0, False)
Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
               line_, C_display_wenti.m_inner_lin(num, 1), _
                line_, C_display_wenti.m_inner_lin(num, 2), _
                 0, 0, False)
Else
Call set_element_depend(line_, C_display_wenti.m_inner_lin(num, 2), _
               point_, C_display_wenti.m_inner_poi(num, 2), _
                circle_, C_display_wenti.m_inner_circ(num, 1), _
                 0, 0, False)
Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
               line_, C_display_wenti.m_inner_lin(num, 2), _
                circle_, C_display_wenti.m_inner_circ(num, 2), _
                 0, 0, False)
End If
End Sub

Private Sub set_initial_condition_31(ByVal num As Integer, temp_record As total_record_type)
'-31 在□□上取一点□使得□□＝□□
Dim tn_%, tl%, tc%, tp%
'inpcond(-31) = 在□□上取一点□使得□□=□□
temp_record.record_data.data0.condition_data.condition_no = 0 ' = record0
  tn_% = 0
  Call set_equal_dline(C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
   C_display_wenti.m_point_no(num, 4), C_display_wenti.m_point_no(num, 5), _
     0, 0, 0, 0, 0, 0, 0, temp_record, tn_%, 0, 0, 0, 0, True)
     Call C_display_wenti.set_m_condition_data(num, eline_, tn_%)
tl% = C_display_wenti.m_inner_lin(num, 1)
tc% = C_display_wenti.m_inner_circ(num, 1)
tp% = C_display_wenti.m_inner_poi(num, 1)
'**********************************************************************************************
Call add_point_to_line(tp%, tl%, 0, False, False, 0)
Call set_element_depend(circle_, tc%, point_, m_Circ(tc%).data(0).data0.center, _
       0, 0, 0, 0, False)
'半径固定的圆
m_Circ(tc%).data(0).radii_depend_poi(0) = C_display_wenti.m_inner_poi(num, 3)
m_Circ(tc%).data(0).radii_depend_poi(1) = C_display_wenti.m_inner_poi(num, 4)
Call set_element_depend(point_, tp%, line_, tl%, circle_, tc%, 0, 0, False)
m_Circ(tc%).data(0).depend_para = 1
If regist_data.run_type = 1 Then
Call set_V_coordinate_system(C_display_wenti.m_point_no(num, 0), _
        C_display_wenti.m_point_no(num, 1))
temp_record.record_data.data0.condition_data.condition_no = 0
Call add_conditions_to_record(eline_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
Call set_equal_v_line(C_display_wenti.m_point_no(num, 2), _
                        C_display_wenti.m_point_no(num, 3), _
                       C_display_wenti.m_point_no(num, 4), _
                        C_display_wenti.m_point_no(num, 5), _
                          temp_record)
End If
End Sub

Private Sub set_initial_condition_30_58(ByVal num As Integer, temp_record As total_record_type)
'-30 在⊙□[down\\(_)]上取一点□使得□□＝□□
'-58 在⊙□□□上取一点□使得□□＝□□
Dim tn_%
Dim tp(3) As Integer
Dim tc(1) As Integer
'inpcond(-30) = 在⊙□(_)上取一点□使得□□=□□
temp_record.record_data.data0.condition_data.condition_no = 0 ' = record0
tn_% = 0
  tp(0) = C_display_wenti.m_inner_poi(num, 1)
  tp(1) = C_display_wenti.m_inner_poi(num, 2)
  tp(2) = C_display_wenti.m_inner_poi(num, 3)
  tp(3) = C_display_wenti.m_inner_poi(num, 4)
  tc(0) = C_display_wenti.m_inner_circ(num, 1)
  tc(1) = C_display_wenti.m_inner_circ(num, 2)
   Call set_element_depend(circle_, tc(1), point_, m_Circ(tc(1)).data(0).data0.center, _
           0, 0, 0, 0, False)
  m_Circ(tc(1)).data(0).radii_depend_poi(0) = tp(2)
  m_Circ(tc(1)).data(0).radii_depend_poi(1) = tp(3)
  m_Circ(tc(1)).data(0).depend_para = 1
  If m_Circ(tc(0)).data(0).data0.in_point(3) = 0 Or _
       m_Circ(tc(0)).data(0).data0.in_point(3) > _
         m_Circ(tc(0)).data(0).data0.center Then
          Call set_element_depend(circle_, tc(0), point_, m_Circ(tc(0)).data(0).data0.center, _
           point_, m_Circ(tc(0)).data(0).data0.in_point(1), 0, 0, False)
  End If
          Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
             circle_, tc(0), circle_, tc(1), 0, 0, False)
  Call set_equal_dline(tp(0), tp(1), tp(2), tp(3), _
    0, 0, 0, 0, 0, 0, 0, temp_record, tn_%, 0, 0, 0, 0, False)
Call C_display_wenti.set_m_condition_data(num, eline_, tn_%)
If m_Circ(tc(0)).data(0).degree = 0 And m_Circ(tc(1)).data(0).degree = 0 Then
     m_poi(C_display_wenti.m_inner_poi(num, 1)).data(0).degree_for_reduce = 0
ElseIf (m_Circ(tc(0)).data(0).degree >= 1 And m_Circ(tc(1)).data(0).degree = 0) Or _
        (m_Circ(tc(0)).data(0).degree = 0 And m_Circ(tc(1)).data(0).degree >= 1) Then
     m_poi(C_display_wenti.m_inner_poi(num, 1)).data(0).degree_for_reduce = 1
ElseIf m_Circ(tc(0)).data(0).degree >= 1 And m_Circ(tc(1)).data(0).degree >= 1 Then
     m_poi(C_display_wenti.m_inner_poi(num, 1)).data(0).degree_for_reduce = 2
End If
If regist_data.run_type = 1 Then
   Call add_new_point_on_circle_for_vector(C_display_wenti.m_point_no(num, 42), tp(0), _
          temp_record)
   temp_record.record_data.data0.condition_data.condition_no = 0
   Call add_conditions_to_record(eline_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
   Call set_equal_v_line(tp(0), tp(1), tp(2), tp(3), temp_record)
End If
End Sub

Private Sub set_initial_condition_24(ByVal num As Integer, temp_record As total_record_type)

Dim cir(1) As Integer
Dim m%, n%, tn_%
'inpcond(-24) = 弧□□=弧□□
'If c_display_wenti.m_point_no(num,0) <> c_display_wenti.m_point_no(num,2) And _
 c_display_wenti.m_point_no(num,1) <> c_display_wenti.m_point_no(num,2) Then
cir(0) = read_circle_from_chord(C_display_wenti.m_point_no(num, 0), _
  C_display_wenti.m_point_no(num, 1), 0)
'Else
cir(1) = read_circle_from_chord(C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 3), 0)
'End If
temp_record.record_data.data0.condition_data.condition_no = 0 ' = record0
m% = arc_no(C_display_wenti.m_point_no(num, 0), cir(0), C_display_wenti.m_point_no(num, 1))
n% = arc_no(C_display_wenti.m_point_no(num, 2), cir(1), C_display_wenti.m_point_no(num, 3))
If m% > 0 And n% > 0 Then
tn_% = 0
Call set_equal_arc(m%, n%, temp_record, tn_%, 0)
Call C_display_wenti.set_m_condition_data(num, equal_arc_, tn_%)
End If

End Sub

Private Sub set_initial_condition_23_22(ByVal num As Integer, temp_record As total_record_type)
'-22 过□点平行□□的直线交□□于□
'-23 过□点垂直□□的直线交□□于□
Dim tn_%, tn%
'inpcond(-23) =过□点垂直□□的直线交□□于□
'inpcond(-22) = 过□点平行□□的直线交□□于□
t_line(0) = C_display_wenti.m_inner_lin(num, 1) '交□□
t_line(1) = C_display_wenti.m_inner_lin(num, 2) '连线
t_line(2) = C_display_wenti.m_inner_lin(num, 3) '平行□□
Call set_element_depend(line_, t_line(1), _
                  line_, t_line(2), point_, _
                   C_display_wenti.m_point_no(num, 0), _
                    0, 0, False)
Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
                  line_, t_line(0), line_, t_line(1), 0, 0, False)                    '交点
If C_display_wenti.m_no(num) = -22 Then
   temp_record.record_data.data0.condition_data.condition_no = 0 ' = record0
   tn_% = 0
   Call set_dparal(t_line(1), t_line(2), temp_record, tn_%, 0, True)
   Call C_display_wenti.set_m_condition_data(num, paral_, tn_%)
ElseIf C_display_wenti.m_no(num) = -23 Then
   temp_record.record_data.data0.condition_data.condition_no = 0 ' = record0
   tn_% = 0
   Call set_dverti(t_line(1), t_line(2), temp_record, tn_%, 0, True)
   Call C_display_wenti.set_m_condition_data(num, verti_, tn_%)
End If
End Sub

Private Sub set_initial_condition_21(ByVal num As Integer, temp_record As total_record_type)
Dim tn_%
'depend_no(num) = 1
temp_record.record_data.data0.condition_data.condition_no = 0 ' = record0
  t_line(0) = line_number0(C_display_wenti.m_point_no(num, 1), _
      C_display_wenti.m_point_no(num, 0), 0, 0)
  t_line(1) = line_number0(C_display_wenti.m_point_no(num, 0), _
      C_display_wenti.m_point_no(num, 2), 0, 0)
tn_% = 0
Call set_dverti(t_line(0), t_line(1), temp_record, tn_%, 0, True)
Call C_display_wenti.set_m_condition_data(num, verti_, tn_%)
'Call draw_picture_21(num, no_reduce)

End Sub

Private Sub set_initial_condition_17_18(ByVal num As Integer, temp_record As total_record_type)
Dim tn_%
Dim tp(2) As Integer
Dim tl(1) As Integer
Dim ang(2) As Integer
tp(0) = C_display_wenti.m_point_no(num, 0)
tp(1) = C_display_wenti.m_point_no(num, 1)
tp(2) = C_display_wenti.m_point_no(num, 2)
If tp(1) > tp(2) Then
  Call exchange_two_integer(tp(1), tp(2))
End If
Call draw_triangle(tp(0), tp(1), tp(2), condition)
tl(0) = line_number0(tp(0), tp(1), 0, 0)
tl(1) = line_number0(tp(0), tp(2), 0, 0)

temp_record.record_data.data0.condition_data.condition_no = 0 ' = record0
If C_display_wenti.m_no(num) = -17 Then
tn_% = 0
 Call set_dverti(tl(0), tl(1), temp_record, tn_%, 0, True)
    Call C_display_wenti.set_m_condition_data(num, verti_, tn_%)
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
 Call set_Drelation(C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
     0, 0, 0, 0, 0, 0, "'2", temp_record, 0, 0, 0, 0, 0, False)
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
 Call set_Drelation(C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), _
  C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 2), _
     0, 0, 0, 0, 0, 0, "'2", temp_record, 0, 0, 0, 0, 0, False)
'For i% = 1 To 2
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
  ang(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 0), _
    C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), 0, 0))
    If ang(0) > 0 Then
     Call set_angle_value(ang(0), "45", temp_record, 0, 0, True)
    End If
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
   ang(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 0), _
    C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 1), 0, 0))
  If ang(0) > 0 Then
   Call set_angle_value(ang(0), "45", temp_record, 0, 0, True)
  End If
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
  ang(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 1), _
    C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 2), 0, 0))
   If ang(0) > 0 Then
    Call set_angle_value(Abs(angle_number(C_display_wenti.m_point_no(num, 1), _
     C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 2), 0, 0)), _
      "90", temp_record, 0, 0, True)
   End If
End If
 '*******************************************************
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
 ang(0) = Abs(angle_number(C_display_wenti.m_point_no(num, 0), _
     C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), 0, 0))
  ang(1) = Abs(angle_number(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 2), _
          C_display_wenti.m_point_no(num, 1), 0, 0))
If ang(0) > 0 And ang(1) > 0 Then
tn_% = 0
 Call set_three_angle_value(ang(0), ang(1), 0, "1", "-1", "0", "0", _
             0, temp_record, tn_%, 0, 0, 0, 0, 0, True)
             Call C_display_wenti.set_m_condition_data(num, angle3_value_, tn_%)
End If
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
tn_% = 0
  Call set_equal_dline(C_display_wenti.m_point_no(num, 0), _
   C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 0), _
      C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, 0, 0, _
       temp_record, tn_%, 0, 0, 0, 0, False)
       Call C_display_wenti.set_m_condition_data(num, eline_, tn_%)

End Sub

Private Sub set_initial_condition_14_13_11_10(ByVal num As Integer, temp_record As total_record_type)
'-10 □□□□是平行四边形
'-11 □□□□是平行四边形
'-13 □□□□是长方形
'-14 □□□□是等腰梯形
Dim tn_%, i%
Dim t_degree As Byte
Dim t_lv(3) As Integer
Dim tn(3) As Integer
Dim para(3) As String
Dim c_data As condition_data_type
If m_poi(C_display_wenti.m_point_no(num, 0)).data(0).degree_for_reduce = 0 And _
    m_poi(C_display_wenti.m_point_no(num, 1)).data(0).degree_for_reduce = 0 Then
     t_degree = 0
ElseIf (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).degree_for_reduce > 0 And _
    m_poi(C_display_wenti.m_point_no(num, 1)).data(0).degree_for_reduce = 0) Or _
       (m_poi(C_display_wenti.m_point_no(num, 0)).data(0).degree_for_reduce = 0 And _
    m_poi(C_display_wenti.m_point_no(num, 1)).data(0).degree_for_reduce > 0) Then
     t_degree = 1
ElseIf m_poi(C_display_wenti.m_point_no(num, 0)).data(0).degree_for_reduce > 0 And _
    m_poi(C_display_wenti.m_point_no(num, 1)).data(0).degree_for_reduce > 0 Then
     t_degree = 2
End If
        Call draw_polygon4(C_display_wenti.m_point_no(num, 0), _
                           C_display_wenti.m_point_no(num, 1), _
                           C_display_wenti.m_point_no(num, 2), _
                           C_display_wenti.m_point_no(num, 3), condition)
If C_display_wenti.m_no(num) = -13 Then
'depend_no(num) = 1
tn_% = 0
Call set_long_squre(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
      C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
       temp_record, tn_%, 0, 0, True)
       Call C_display_wenti.set_m_condition_data(num, long_squre_, tn_%)
       If t_degree = 0 Then
          m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce = _
            min_for_byte(1, m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce)
       ElseIf t_degree > 0 Then
          m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce = _
            min_for_byte(1, m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce)
       End If
          m_poi(C_display_wenti.m_point_no(num, 3)).data(0).degree_for_reduce = _
            min_for_byte(1, m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce)
If regist_data.run_type = 1 Then
    For i% = 0 To 3
    If last_conditions.last_cond(0).v_line_value_no < 2 Then
     Call set_V_coordinate_system(C_display_wenti.m_point_no(num, i%), _
             C_display_wenti.m_point_no(num, (i% + 1) Mod 3))
    End If
    Next i%
    Call add_conditions_to_record(long_squre_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
    t_lv(0) = vector_number(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_point_no(num, 1), "")
    t_lv(1) = vector_number(C_display_wenti.m_point_no(num, 3), _
             C_display_wenti.m_point_no(num, 2), "")
      Call set_item0(t_lv(0), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), _
            0, c_data, 0, tn(0), 0, 0, temp_record.record_data.data0.condition_data, False)
      Call set_item0(t_lv(1), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), _
            0, c_data, 0, tn(1), 0, 0, temp_record.record_data.data0.condition_data, False)
     Call set_general_string(tn(0), tn(1), 0, 0, para(0), time_string("-1", para(1), True, False), _
           "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
    t_lv(2) = vector_number(C_display_wenti.m_point_no(num, 1), _
             C_display_wenti.m_point_no(num, 2), "")
    t_lv(3) = vector_number(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_point_no(num, 3), "")
       temp_record.record_data.data0.condition_data.condition_no = 0
    Call add_conditions_to_record(long_squre_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
      Call set_item0(t_lv(0), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), _
            0, c_data, 0, tn(0), 0, 0, temp_record.record_data.data0.condition_data, False)
      Call set_item0(t_lv(1), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), _
            0, c_data, 0, tn(1), 0, 0, temp_record.record_data.data0.condition_data, False)
      Call set_general_string(tn(0), tn(1), 0, 0, para(0), time_string("-1", para(1), True, False), _
           "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
 '**************************************************************************************************
      Call set_item0(t_lv(0), t_lv(2), 0, 0, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "0", para(0), _
            0, temp_record.record_data.data0.condition_data, 0, tn(0), 0, 0, c_data, False)
      Call set_item0(t_lv(0), t_lv(3), 0, 0, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "0", para(0), _
            0, temp_record.record_data.data0.condition_data, 0, tn(0), 0, 0, c_data, False)
      Call set_item0(t_lv(1), t_lv(2), 0, 0, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "0", para(0), _
            0, temp_record.record_data.data0.condition_data, 0, tn(0), 0, 0, c_data, False)
      Call set_item0(t_lv(1), t_lv(3), 0, 0, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "0", para(0), _
            0, temp_record.record_data.data0.condition_data, 0, tn(0), 0, 0, c_data, False)
 End If
ElseIf C_display_wenti.m_no(num) = -11 Then
'depend_no(num) = 1
tn_% = 0
Call set_parallelogram(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
      C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
        temp_record, tn_%, 0)
         't_lv(0) = line_number(Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(0), _
                             Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(1), _
                             condition, True, 0, _
                             point_, Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(0), _
                             point_, Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(1), _
                             0, True)
         't_lv(1) = line_number(Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(2), _
                             Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(3), _
                             condition, True, 0, _
                             point_, Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(2), _
                             line_, t_lv(0), 0, True)
         'Call line_number(Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(1), _
                             Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(2), _
                             condition, True, 0, _
                             point_, Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(1), _
                             point_, Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(2), _
                             0, True)
         'Call line_number(Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(0), _
                             Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(3), _
                             condition, True, 0, _
                             point_, Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(0), _
                             point_, Dpolygon4(Dparallelogram(tn_%).data(0).polygon4_no).data(0).poi(3), _
                             0, True)
 '******************************************************************************************************************
        Call C_display_wenti.set_m_condition_data(num, parallelogram_, tn_%)
        If t_degree = 0 And m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce = 0 Then
           m_poi(C_display_wenti.m_point_no(num, 3)).data(0).degree_for_reduce = 0
        ElseIf (t_degree > 0 And _
            m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce = 0) Or _
               (t_degree = 0 And _
            m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce > 0) Then
           m_poi(C_display_wenti.m_point_no(num, 3)).data(0).degree_for_reduce = _
              min_for_byte(1, _
                 m_poi(C_display_wenti.m_point_no(num, 3)).data(0).degree_for_reduce)
        ElseIf t_degree > 0 And _
            m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce > 0 Then
           m_poi(C_display_wenti.m_point_no(num, 3)).data(0).degree_for_reduce = _
              min_for_byte(2, _
                 m_poi(C_display_wenti.m_point_no(num, 3)).data(0).degree_for_reduce)
        End If
           m_poi(C_display_wenti.m_point_no(num, 3)).data(0).degree = 0
 If regist_data.run_type = 1 Then
    For i% = 0 To 3
    If last_conditions.last_cond(0).v_line_value_no < 2 Then
     Call set_V_coordinate_system(C_display_wenti.m_point_no(num, i%), _
             C_display_wenti.m_point_no(num, (i% + 1) Mod 3))
    End If
    Next i%
    Call add_conditions_to_record(parallelogram_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
    t_lv(0) = vector_number(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_point_no(num, 1), "")
    t_lv(1) = vector_number(C_display_wenti.m_point_no(num, 3), _
             C_display_wenti.m_point_no(num, 2), "")
      Call set_item0(t_lv(0), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), _
            0, c_data, 0, tn(0), 0, 0, temp_record.record_data.data0.condition_data, False)
      Call set_item0(t_lv(1), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), _
            0, c_data, 0, tn(1), 0, 0, temp_record.record_data.data0.condition_data, False)
     Call set_general_string(tn(0), tn(1), 0, 0, para(0), time_string("-1", para(1), True, False), _
           "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
           temp_record.record_data.data0.condition_data.condition_no = 0
     Call add_conditions_to_record(parallelogram_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
    t_lv(0) = vector_number(C_display_wenti.m_point_no(num, 1), _
             C_display_wenti.m_point_no(num, 2), "")
    t_lv(1) = vector_number(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_point_no(num, 3), "")
      Call set_item0(t_lv(0), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), _
            0, c_data, 0, tn(0), 0, 0, temp_record.record_data.data0.condition_data, False)
      Call set_item0(t_lv(1), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), _
            0, c_data, 0, tn(1), 0, 0, temp_record.record_data.data0.condition_data, False)
      Call set_general_string(tn(0), tn(1), 0, 0, para(0), time_string("-1", para(1), True, False), _
           "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
 End If
ElseIf C_display_wenti.m_no(num) = -14 Then
'depend_no(num) = 1
tn_% = 0
Call set_tixing(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
      C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), equal_side_tixing_, _
       temp_record, 0, 0)
       Call C_display_wenti.set_m_condition_data(num, tixing_, tn_%)
        If t_degree = 0 And m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce = 0 Then
           m_poi(C_display_wenti.m_point_no(num, 3)).data(0).degree_for_reduce = 0
        ElseIf (t_degree > 0 And _
            m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce = 0) Or _
               (t_degree = 0 And _
            m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce > 0) Then
           m_poi(C_display_wenti.m_point_no(num, 3)).data(0).degree_for_reduce = _
              min_for_byte(1, _
                 m_poi(C_display_wenti.m_point_no(num, 3)).data(0).degree_for_reduce)
        ElseIf t_degree > 0 And _
            m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce > 0 Then
           m_poi(C_display_wenti.m_point_no(num, 3)).data(0).degree_for_reduce = _
              min_for_byte(2, _
                 m_poi(C_display_wenti.m_point_no(num, 3)).data(0).degree_for_reduce)
        End If
 If regist_data.run_type = 1 Then
    For i% = 0 To 3
    If last_conditions.last_cond(0).v_line_value_no < 2 Then
     Call set_V_coordinate_system(C_display_wenti.m_point_no(num, i%), _
             C_display_wenti.m_point_no(num, (i% + 1) Mod 3))
    End If
    Next i%
    t_lv(0) = vector_number(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_point_no(num, 1), "")
    t_lv(1) = vector_number(C_display_wenti.m_point_no(num, 3), _
             C_display_wenti.m_point_no(num, 2), "")
      Call set_item0(t_lv(0), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), _
            0, c_data, 0, tn(0), 0, 0, c_data, False)
      Call set_item0(t_lv(1), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), _
            0, c_data, 0, tn(1), 0, 0, c_data, False)
      temp_record.record_data.data0.condition_data.condition(8).ty = new_point_
        Call set_general_string(tn(0), tn(1), 0, 0, para(0), time_string("-a", para(1), True, False), _
           "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
 End If
Else
'depend_no(num) = 1
Call set_rhombus(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
      C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
       temp_record, tn_%, 0)
       
       Call C_display_wenti.set_m_condition_data(num, tixing_, tn_%)
       If t_degree = 0 Then
          m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce = _
            min_for_byte(1, m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce)
       ElseIf t_degree > 0 Then
          m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce = _
            min_for_byte(1, m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce)
       End If
          m_poi(C_display_wenti.m_point_no(num, 3)).data(0).degree_for_reduce = _
            min_for_byte(1, m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce)
 If regist_data.run_type = 1 Then
    For i% = 0 To 3
    If last_conditions.last_cond(0).v_line_value_no < 2 Then
     Call set_V_coordinate_system(C_display_wenti.m_point_no(num, i%), _
             C_display_wenti.m_point_no(num, (i% + 1) Mod 3))
    End If
    Next i%
    t_lv(0) = vector_number(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_point_no(num, 1), "")
    t_lv(1) = vector_number(C_display_wenti.m_point_no(num, 3), _
             C_display_wenti.m_point_no(num, 2), "")
       temp_record.record_data.data0.condition_data.condition_no = 0
      Call add_conditions_to_record(rhombus_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
      Call set_item0(t_lv(0), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), _
            0, c_data, 0, tn(0), 0, 0, temp_record.record_data.data0.condition_data, False)
      Call set_item0(t_lv(1), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), _
            0, c_data, 0, tn(1), 0, 0, temp_record.record_data.data0.condition_data, False)
     Call set_general_string(tn(0), tn(1), 0, 0, para(0), time_string("-1", para(1), True, False), _
           "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
    t_lv(2) = vector_number(C_display_wenti.m_point_no(num, 1), _
             C_display_wenti.m_point_no(num, 2), "")
    t_lv(3) = vector_number(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_point_no(num, 3), "")
       temp_record.record_data.data0.condition_data.condition_no = 0
      Call add_conditions_to_record(rhombus_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
      Call set_item0(t_lv(2), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), _
            0, c_data, 0, tn(0), 0, 0, temp_record.record_data.data0.condition_data, False)
      Call set_item0(t_lv(1), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), _
            0, c_data, 0, tn(3), 0, 0, temp_record.record_data.data0.condition_data, False)
      Call set_general_string(tn(0), tn(1), 0, 0, para(0), time_string("-1", para(1), True, False), _
           "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
       temp_record.record_data.data0.condition_data.condition_no = 0
      Call add_conditions_to_record(rhombus_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
      Call set_item0(t_lv(0), -10, t_lv(0), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), _
            0, c_data, 0, tn(0), 0, 0, temp_record.record_data.data0.condition_data, False)
      Call set_item0(t_lv(1), -10, t_lv(1), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), _
            0, c_data, 0, tn(1), 0, 0, temp_record.record_data.data0.condition_data, False)
      Call set_item0(t_lv(2), -10, t_lv(2), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(2), _
            0, c_data, 0, tn(2), 0, 0, temp_record.record_data.data0.condition_data, True)
      Call set_item0(t_lv(3), -10, t_lv(3), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(3), _
            0, c_data, 0, tn(3), 0, 0, temp_record.record_data.data0.condition_data, True)
      If tn(0) = 0 Then
         Call set_item0_value(tn(1), 0, 0, 0, "", "", para(0), 0, temp_record.record_data.data0.condition_data)
         Call set_item0_value(tn(2), 0, 0, 0, "", "", para(0), 0, temp_record.record_data.data0.condition_data)
         Call set_item0_value(tn(3), 0, 0, 0, "", "", para(0), 0, temp_record.record_data.data0.condition_data)
      ElseIf tn(1) = 0 Then
         Call set_item0_value(tn(0), 0, 0, 0, "", "", para(1), 0, temp_record.record_data.data0.condition_data)
         Call set_item0_value(tn(2), 0, 0, 0, "", "", para(1), 0, temp_record.record_data.data0.condition_data)
         Call set_item0_value(tn(3), 0, 0, 0, "", "", para(1), 0, temp_record.record_data.data0.condition_data)
      ElseIf tn(2) = 0 Then
         Call set_item0_value(tn(0), 0, 0, 0, "", "", para(2), 0, temp_record.record_data.data0.condition_data)
         Call set_item0_value(tn(1), 0, 0, 0, "", "", para(2), 0, temp_record.record_data.data0.condition_data)
         Call set_item0_value(tn(3), 0, 0, 0, "", "", para(2), 0, temp_record.record_data.data0.condition_data)
      ElseIf tn(3) = 0 Then
         Call set_item0_value(tn(0), 0, 0, 0, "", "", para(3), 0, temp_record.record_data.data0.condition_data)
         Call set_item0_value(tn(1), 0, 0, 0, "", "", para(3), 0, temp_record.record_data.data0.condition_data)
         Call set_item0_value(tn(2), 0, 0, 0, "", "", para(3), 0, temp_record.record_data.data0.condition_data)
      Else
      Call set_general_string(tn(0), tn(1), 0, 0, para(0), time_string("-1", para(1), True, False), _
           "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
      Call set_general_string(tn(0), tn(2), 0, 0, para(0), time_string("-1", para(2), True, False), _
           "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
      Call set_general_string(tn(0), tn(3), 0, 0, para(0), time_string("-1", para(3), True, False), _
           "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
      Call set_general_string(tn(1), tn(2), 0, 0, para(1), time_string("-1", para(2), True, False), _
           "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
      Call set_general_string(tn(1), tn(3), 0, 0, para(1), time_string("-1", para(3), True, False), _
           "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
      Call set_general_string(tn(2), tn(3), 0, 0, para(2), time_string("-1", para(3), True, False), _
           "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
      End If
  End If
End If
End Sub

Private Sub set_initial_condition_5(ByVal num As Integer, temp_record As total_record_type)
Dim ang As Integer
Dim tn_%
Dim value As String
ang = angle_number(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), 0, 0)
 value = initial_string(number_string(C_display_wenti.m_point_no(num, 3))) 'initial_string(cond_to_string(num, 3, 18, 0))
       If InStr(1, value, ".", 0) > 0 Then
        th_chose(-5).chose = 2
      End If
If ang <> 0 Then
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
tn_% = 0
Call set_angle_value(Abs(ang), value, temp_record, tn_%, 0, True)
Call C_display_wenti.set_m_condition_data(num, angle3_value_, tn_%)
End If

End Sub

Private Sub set_initial_condition_4(ByVal num As Integer, temp_record As total_record_type)
'∠□□□=∠□□□
Dim ang(1) As Integer
Dim tn_%
If num > 0 Then
   If C_display_wenti.m_inner_poi(num, 1) = 0 Then
'   ***********************************************************************************
Call set_wenti_cond_4(C_display_wenti.m_point_no(num, 0), _
                       C_display_wenti.m_point_no(num, 1), _
                        C_display_wenti.m_point_no(num, 2), _
                      C_display_wenti.m_point_no(num, 3), _
                       C_display_wenti.m_point_no(num, 4), _
                        C_display_wenti.m_point_no(num, 5), _
                         num)
'*************************************************************************************
   End If
End If
If C_display_wenti.m_no_(num) = 50 Then
 Call set_initial_condition_50(num, temp_record)
Else
ang(0) = angle_number(C_display_wenti.m_point_no(num, 0), _
 C_display_wenti.m_point_no(num, 1), C_display_wenti.m_point_no(num, 2), 0, 0)
ang(1) = angle_number(C_display_wenti.m_point_no(num, 3), _
 C_display_wenti.m_point_no(num, 4), C_display_wenti.m_point_no(num, 5), 0, 0)
  If ang(0) <> 0 And ang(1) <> 0 Then
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
tn_% = 0
Call set_three_angle_value(Abs(ang(0)), Abs(ang(1)), 0, _
       "1", "-1", "0", "0", 0, temp_record, tn_%, 0, 0, 0, 0, 0, True)
       Call C_display_wenti.set_m_condition_data(num, angle3_value_, tn_%)
End If
End If
End Sub

Private Sub set_initial_condition_1(ByVal num As Integer, temp_record As total_record_type)
Dim tn_%, tc%, tl%, tc2%, no_%
   temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
   tn_% = 0
   If C_display_wenti.m_inner_poi(num, 1) = 0 Then
      Call set_wenti_cond_1(C_display_wenti.m_point_no(num, 0), _
                             C_display_wenti.m_point_no(num, 1), _
                              C_display_wenti.m_point_no(num, 2), _
                               C_display_wenti.m_point_no(num, 3), _
                                 num, line_, C_display_wenti.m_inner_lin(num, 1), _
                                  circle_, C_display_wenti.m_inner_circ(num, 1), _
                                    C_display_wenti.m_inner_point_type(num), _
                                      C_display_wenti.m_inner_poi(num, 1))
   End If
   tc% = C_display_wenti.m_inner_circ(num, 1)
   tl% = C_display_wenti.m_inner_lin(num, 1)
   tc2% = C_display_wenti.m_inner_circ(num, 2)
   no_% = C_display_wenti.m_no_(num)
   If tc% > 0 Then
    m_Circ(tc2%).data(0).radii_depend_poi(0) = C_display_wenti.m_inner_poi(num, 3)
    m_Circ(tc2%).data(0).radii_depend_poi(1) = C_display_wenti.m_inner_poi(num, 4)
    m_Circ(tc2%).data(0).depend_para = 1
      Call set_element_depend(circle_, tc2, point_, m_Circ(tc2%).data(0).data0.center, _
                               0, 0, 0, 0, False)
   End If
   If no_% = 4 Then '-54 □□的垂直平分线交□□于□
      Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
                point_, C_display_wenti.m_inner_poi(num, 3), _
                 point_, C_display_wenti.m_inner_poi(num, 4), 0, 0, False)
   ElseIf no_% = -54 Then '-54 □□的垂直平分线交□□于□
      Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
                point_, C_display_wenti.m_inner_poi(num, 3), _
                 point_, C_display_wenti.m_inner_poi(num, 4), _
                  line_, tl%, True)
   ElseIf no_% = -53 Then '-53 □□的垂直平分线交⊙□[down\\(_)]于□
      Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
                point_, C_display_wenti.m_inner_poi(num, 3), _
                 point_, C_display_wenti.m_inner_poi(num, 4), _
                  circle_, tc%, True)
   ElseIf no_% = -30 Then '-30 在⊙□[down\\(_)]上取一点□使得□□＝□□
           Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
                circle_, tc2%, circle_, tc%, 0, 0, True)
   ElseIf no_% = -31 Then '-31 在□□上取一点□使得□□＝□□
           Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
                circle_, tc2%, circle_, tl%, 0, 0, True)
   Else
           Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
                circle_, tc2%, point_, C_display_wenti.m_inner_poi(num, 1), 0, 0, False)
   End If
   Call set_equal_dline(C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 1), _
   C_display_wenti.m_point_no(num, 2), C_display_wenti.m_point_no(num, 3), _
     0, 0, 0, 0, 0, 0, 0, temp_record, tn_%, 0, 0, 0, 0, True)
   'Call set_initial_data_for_draw_1(num)
    'Call C_display_wenti.set_m_condition_data(num,angle3_value_, tn_%)

End Sub
Private Sub set_initial_condition_2(ByVal num As Integer, temp_record As total_record_type)
'-2 作⊙□[down\\(_)]和⊙□[down\\(_)]的公切线□□
'-60 作⊙□□□和⊙□□□的公切线□□
'-59 作⊙□□□和⊙□[down\\(_)]的公切线□□
Dim k%, c1%, c2%
c1% = C_display_wenti.m_inner_circ(num, 1)
c2% = C_display_wenti.m_inner_circ(num, 2)
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
  set_initial_data_for_draw_2 (num)
 If m_Circ(c1%).data(0).data0.in_point(3) = 0 Or _
     m_Circ(c1%).data(0).data0.in_point(3) > m_Circ(c1%).data(0).data0.center Then
  Call set_element_depend(circle_, c1%, _
            point_, m_Circ(c1%).data(0).data0.center, _
             point_, m_Circ(c1%).data(0).data0.in_point(1), _
              0, 0, False)
 End If
 If m_Circ(c2%).data(0).data0.in_point(3) = 0 Or _
     m_Circ(c2%).data(0).data0.in_point(3) > m_Circ(c2%).data(0).data0.center Then
   Call set_element_depend(circle_, c2%, _
            point_, m_Circ(c2%).data(0).data0.center, _
             point_, m_Circ(c2%).data(0).data0.in_point(1), _
              0, 0, False)
 End If
  Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
            circle_, c1%, point_, c2%, 0, 0, True)
  Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 2), _
            circle_, c1%, point_, c2%, 0, 0, True)
  If C_display_wenti.m_inner_poi(num, 1) = _
        C_display_wenti.m_inner_poi(num, 2) Then
   Call set_element_depend(line_, C_display_wenti.m_inner_lin(num, 1), _
            point_, C_display_wenti.m_inner_poi(num, 1), _
             line_, C_display_wenti.m_inner_lin(num, 2), _
              0, 0, False)
  Else
   Call set_element_depend(line_, C_display_wenti.m_inner_lin(num, 1), _
            point_, C_display_wenti.m_inner_poi(num, 1), _
             point_, C_display_wenti.m_inner_poi(num, 2), _
              0, 0, False)
  End If
'   Call set_tangent_line(C_display_wenti.m_inner_lin(num, 1), _
           C_display_wenti.m_inner_poi(num, 1), _
            C_display_wenti.m_inner_circ(num, 1), _
             C_display_wenti.m_inner_poi(num, 2), _
              C_display_wenti.m_inner_circ(num, 2), _
               temp_record, 0, 0)
End Sub

Private Sub set_initial_condition1(ByVal num As Integer, temp_record As total_record_type)
'1 直线□□上任取一点□
'depend_no(num) = 1
Dim c_data0 As condition_data_type
t_line(0) = C_display_wenti.m_inner_lin(num, 1)
Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
          line_, t_line(0), 0, 0, 0, 0, False)
End Sub



Private Sub set_initial_condition5_15(ByVal num As Integer, temp_record As total_record_type)
'5 取线段□□的中点□
'15以□□为直径作⊙□[down\\(_)]
Dim tn_%, n%
Dim c_data(2) As condition_data_type
Dim t_lv(1) As Integer
Dim tn(1) As Integer
Dim lv As V_line_value_data0_type
Dim para(1) As String
Dim tv As String
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
   t_line(0) = C_display_wenti.m_inner_lin(num, 1)
   Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), point_, _
         C_display_wenti.m_point_no(num, 0), point_, _
          C_display_wenti.m_point_no(num, 1), 0, 0, True) '设置新点与已有条件的相关性
   m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree = 0
   
If C_display_wenti.m_no(num) = 15 Then '园的直径
'Call add_point_to_line(C_display_wenti.m_point_no(num,2), _
    t_line(0), 0, no_display, False, 0, c_data(0))
     n% = C_display_wenti.m_inner_circ(num, 1)
     m_Circ(n%).data(0).last_Diameter = m_Circ(n%).data(0).last_Diameter
     m_Circ(n%).data(0).Diameter(m_Circ(n%).data(0).last_Diameter).poi(0) = _
                C_display_wenti.m_point_no(num, 0)
     m_Circ(n%).data(0).Diameter(m_Circ(n%).data(0).last_Diameter).poi(1) = _
                C_display_wenti.m_point_no(num, 1)
     If m_Circ(n%).data(0).Diameter(m_Circ(n%).data(0).last_Diameter).poi(0) > _
               m_Circ(n%).data(0).Diameter(m_Circ(n%).data(0).last_Diameter).poi(1) Then
        Call exchange_two_integer(m_Circ(n%).data(0).Diameter(m_Circ(n%).data(0).last_Diameter).poi(0), _
               m_Circ(n%).data(0).Diameter(m_Circ(n%).data(0).last_Diameter).poi(1))
     End If
End If
tn_% = 0
Call set_mid_point(C_display_wenti.m_point_no(num, 0), _
                    C_display_wenti.m_point_no(num, 2), _
                     C_display_wenti.m_point_no(num, 1), _
                      0, 0, 0, 0, 0, temp_record, tn_%, 0, 0, 0, 0)
Call C_display_wenti.set_m_condition_data(num, midpoint_, tn_%)
If regist_data.run_type = 1 Then
 Call add_conditions_to_record(midpoint_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
 c_data(2) = temp_record.record_data.data0.condition_data
   If m_lin(t_line(0)).data(0).parent.element(0).ty = point_ And _
                 m_lin(t_line(0)).data(0).parent.element(1).ty = point_ Then
     Call set_V_coordinate_system( _
      m_lin(t_line(0)).data(0).parent.element(0).no, m_lin(t_line(0)).data(0).parent.element(1).no)
   ElseIf m_lin(t_line(0)).data(0).parent.element(0).ty = point_ Then
     Call set_V_coordinate_system( _
      m_lin(t_line(0)).data(0).parent.element(0).no, C_display_wenti.m_point_no(num, 2))
   ElseIf m_lin(t_line(0)).data(0).parent.element(1).ty = point_ Then
     Call set_V_coordinate_system( _
      C_display_wenti.m_point_no(num, 2), m_lin(t_line(0)).data(0).parent.element(1).no)
   End If
    If is_V_line_value(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_point_no(num, 1), 0, 0, 0, _
              tv$, n%, -1000, 0, 0, 0, lv, False) Then
         temp_record.record_data.data0.condition_data.condition_no = 0
         Call add_conditions_to_record(midpoint_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
         Call add_conditions_to_record(V_line_value_, n%, 0, 0, temp_record.record_data.data0.condition_data)
         Call set_V_line_value(C_display_wenti.m_point_no(num, 0), _
                  C_display_wenti.m_point_no(num, 2), 0, 0, 0, _
                   divide_string(tv$, "2", True, False), temp_record, 0, False)
         Call set_V_line_value(C_display_wenti.m_point_no(num, 2), _
                  C_display_wenti.m_point_no(num, 1), 0, 0, 0, _
                   divide_string(tv$, "2", True, False), temp_record, 0, False)
    ElseIf is_V_line_value(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_point_no(num, 2), 0, 0, 0, _
              tv$, n%, -1000, 0, 0, 0, lv, False) Then
         temp_record.record_data.data0.condition_data.condition_no = 0
         Call add_conditions_to_record(midpoint_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
         Call add_conditions_to_record(V_line_value_, n%, 0, 0, temp_record.record_data.data0.condition_data)
         Call set_V_line_value(C_display_wenti.m_point_no(num, 0), _
                  C_display_wenti.m_point_no(num, 1), 0, 0, 0, _
                   time_string(tv$, "2", True, False), temp_record, 0, False)
         Call set_V_line_value(C_display_wenti.m_point_no(num, 2), _
                  C_display_wenti.m_point_no(num, 1), 0, 0, 0, _
                    tv$, temp_record, 0, False)
    ElseIf is_V_line_value(C_display_wenti.m_point_no(num, 2), _
             C_display_wenti.m_point_no(num, 1), 0, 0, 0, _
              tv$, n%, -1000, 0, 0, 0, lv, False) Then
         temp_record.record_data.data0.condition_data.condition_no = 0
         Call add_conditions_to_record(midpoint_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
         Call add_conditions_to_record(V_line_value_, n%, 0, 0, temp_record.record_data.data0.condition_data)
         Call set_V_line_value(C_display_wenti.m_point_no(num, 0), _
                  C_display_wenti.m_point_no(num, 1), 0, 0, 0, _
                   time_string(tv$, "2", True, False), temp_record, 0, False)
         Call set_V_line_value(C_display_wenti.m_point_no(num, 0), _
                  C_display_wenti.m_point_no(num, 2), 0, 0, 0, _
                    tv$, temp_record, 0, False)
    Else
    t_lv(0) = vector_number(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_point_no(num, 1), "")
    t_lv(1) = vector_number(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_point_no(num, 2), "")
         c_data(1).condition_no = 0
    Call set_item0(t_lv(0), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", _
          para(0), 0, c_data(0), 0, tn(0), 0, 0, c_data(1), False)
    Call set_item0(t_lv(1), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", _
          para(1), 0, c_data(0), 0, tn(1), 0, 0, c_data(1), False)
       Call add_record_to_record(c_data(1), temp_record.record_data.data0.condition_data)
    Call set_general_string(tn(0), tn(1), 0, 0, para(0), time_string("-2", para(1), True, False), _
          "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
    temp_record.record_data.data0.condition_data = c_data(2)
    t_lv(0) = vector_number(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_point_no(num, 1), "")
    t_lv(1) = vector_number(C_display_wenti.m_point_no(num, 2), _
             C_display_wenti.m_point_no(num, 1), "")
        c_data(1).condition_no = 0
    Call set_item0(t_lv(0), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", _
          para(0), 0, c_data(0), 0, tn(0), 0, 0, c_data(1), False)
    Call set_item0(t_lv(1), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", _
          para(1), 0, c_data(1), 0, tn(1), 0, 0, c_data(1), False)
       Call add_record_to_record(c_data(1), temp_record.record_data.data0.condition_data)
    Call set_general_string(tn(0), tn(1), 0, 0, para(0), time_string("-2", para(1), True, False), _
          "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
    End If
End If
End Sub

Private Sub set_initial_condition6(ByVal num As Integer, temp_record As total_record_type)
'□是线段□□上分比为!_~的分点
Dim value As Single
Dim tn_%, n%
Dim tn(1) As Integer
Dim t_lv(1) As Integer
Dim para(1) As String
Dim v$
Dim lv As V_line_value_data0_type
Dim c_data0 As condition_data_type
value = value_string(number_string(C_display_wenti.m_point_no(num, 3)))
      If InStr(1, number_string(C_display_wenti.m_point_no(num, 3)), ".", 0) > 0 Then
        th_chose(-5).chose = 2
      End If
      value = 1 / value + 1
      value_for_draw(C_display_wenti.m_point_no(num, 3)) = value
If m_poi(C_display_wenti.m_point_no(num, 0)).data(0).data0.visible = 0 Then
  t_coord1 = minus_POINTAPI( _
            m_poi(C_display_wenti.m_point_no(num, 2)).data(0).data0.coordinate, _
             m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate)
  t_coord1 = divide_POINTAPI_by_number(t_coord1, value)
  t_coord = add_POINTAPI(m_poi(C_display_wenti.m_point_no(num, 1)).data(0).data0.coordinate, t_coord1)
  Call set_point_coordinate(C_display_wenti.m_point_no(num, 0), t_coord, False)
   Call set_point_visible(C_display_wenti.m_point_no(num, 0), 1, False)
End If
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
   t_line(0) = line_number(C_display_wenti.m_point_no(num, 1), _
                           C_display_wenti.m_point_no(num, 2), _
                           pointapi0, pointapi0, _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 1)), _
                           depend_condition(point_, C_display_wenti.m_point_no(num, 2)), _
                           condition, condition_color, 1, 0)
     Call add_point_to_line(C_display_wenti.m_point_no(num, 0), t_line(0), 0, False, False, 0)
'     Call set_parent(line_, t_line(0), point_, C_display_wenti.m_point_no(num, 0), 0, C_display_wenti.m_point_no(num, 2), _
            C_display_wenti.m_point_no(num, 1))
tn_% = 0
Call set_Drelation(C_display_wenti.m_point_no(num, 1), _
 C_display_wenti.m_point_no(num, 0), C_display_wenti.m_point_no(num, 0), _
  C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, 0, number_string(C_display_wenti.m_point_no(num, 3)), temp_record, tn_%, 0, 0, 0, 0, True)
   Call C_display_wenti.set_m_condition_data(num, midpoint_, tn_%)
                   
 If regist_data.run_type = 1 Then
     If m_lin(t_line(0)).data(0).parent.element(1).ty = point_ And _
           m_lin(t_line(0)).data(0).parent.element(2).ty = point_ Then
     Call set_V_coordinate_system(m_lin(t_line(0)).data(0).parent.element(1).no, _
                                    m_lin(t_line(0)).data(0).parent.element(2).no)
     Else
     Call set_V_coordinate_system(C_display_wenti.m_point_no(num, 1), _
                                    C_display_wenti.m_point_no(num, 2))
     End If
    Call add_conditions_to_record(relation_, tn_%, 0, 0, _
                temp_record.record_data.data0.condition_data)
  If is_V_line_value(C_display_wenti.m_point_no(num, 1), _
          C_display_wenti.m_point_no(num, 2), 0, 0, 0, v$, n%, _
            -1000, 0, 0, 0, lv, False) Then
    Call add_conditions_to_record(V_line_value_, n%, 0, 0, _
                temp_record.record_data.data0.condition_data)
    Call set_V_line_value(C_display_wenti.m_point_no(num, 1), _
           C_display_wenti.m_point_no(num, 0), 0, 0, 0, _
            time_string(v$, divide_string(value, add_string("1", value, False, False), _
              False, False), True, False), temp_record, 0, False)
    Call set_V_line_value(C_display_wenti.m_point_no(num, 0), _
           C_display_wenti.m_point_no(num, 2), 0, 0, 0, _
             divide_string(v$, add_string("1", value, False, False), _
                True, False), temp_record, 0, False)
  ElseIf is_V_line_value(C_display_wenti.m_point_no(num, 1), _
          C_display_wenti.m_point_no(num, 0), 0, 0, 0, v$, n%, _
            -1000, 0, 0, 0, lv, False) Then
    Call add_conditions_to_record(V_line_value_, n%, 0, 0, _
                temp_record.record_data.data0.condition_data)
    Call set_V_line_value(C_display_wenti.m_point_no(num, 0), _
           C_display_wenti.m_point_no(num, 2), 0, 0, 0, _
              divide_string(v$, value, True, False), temp_record, 0, False)
    Call set_V_line_value(C_display_wenti.m_point_no(num, 1), _
           C_display_wenti.m_point_no(num, 2), 0, 0, 0, _
            divide_string(v$, divide_string(value, add_string("1", value, False, False), _
              False, False), True, False), temp_record, 0, False)
  ElseIf is_V_line_value(C_display_wenti.m_point_no(num, 0), _
          C_display_wenti.m_point_no(num, 2), 0, 0, 0, v$, n%, _
            -1000, 0, 0, 0, lv, False) Then
    Call add_conditions_to_record(V_line_value_, n%, 0, 0, _
                temp_record.record_data.data0.condition_data)
    Call set_V_line_value(C_display_wenti.m_point_no(num, 1), _
           C_display_wenti.m_point_no(num, 0), 0, 0, 0, _
              time_string(v$, value, True, False), temp_record, 0, False)
    Call set_V_line_value(C_display_wenti.m_point_no(num, 0), _
           C_display_wenti.m_point_no(num, 2), 0, 0, 0, _
             divide_string(v$, add_string("1", value, False, False), _
                True, False), temp_record, 0, False)
  Else
    t_lv(0) = vector_number(C_display_wenti.m_point_no(num, 1), _
             C_display_wenti.m_point_no(num, 2), "")
    t_lv(1) = vector_number(C_display_wenti.m_point_no(num, 0), _
             C_display_wenti.m_point_no(num, 2), "")
    Call set_item0(t_lv(0), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", _
          para(0), 0, c_data0, 0, tn(0), 0, 0, temp_record.record_data.data0.condition_data, False)
    Call set_item0(t_lv(1), -10, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", _
          para(1), 0, c_data0, 0, tn(1), 0, 0, temp_record.record_data.data0.condition_data, False)
    Call set_general_string(tn(0), tn(1), 0, 0, para(0), _
          time_string(add_string(value, "1", False, False), _
           time_string("-1", para(1), False, False), True, False), _
            "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
  End If
 End If
End Sub
Private Sub set_initial_condition7_61(ByVal num As Integer, temp_record As total_record_type)
'7 ⊙□[down\\(_)]上任取一点□
'depend_no(num) = 1
'Call draw_picture7(num, no_reduce)
Dim i%, l%, tp%, tc%
tp% = C_display_wenti.m_inner_poi(num, 1)
tc% = C_display_wenti.m_inner_circ(num, 1)
Call set_element_depend(point_, tp%, _
                         circle_, tc%, 0, 0, 0, 0, False)
 If m_Circ(tc%).data(0).degree = 0 Then
  m_poi(tp%).data(0).degree_for_reduce = 1
 Else
  m_poi(tp%).data(0).degree_for_reduce = 2
 End If
If regist_data.run_type = 1 Then
 Call add_new_point_on_circle_for_vector(tp%, _
          tc%, temp_record)
End If
End Sub

Private Sub set_initial_condition8_71(ByVal num As Integer, temp_record As total_record_type)
'8 过点□、□、□作⊙
Dim tn_%, tc%
'Call draw_picture8(num, no_reduce)
 If C_display_wenti.m_inner_circ(num, 1) = 0 Then
    tc% = m_circle_number(1, 0, pointapi0, _
            C_display_wenti.m_point_no(num, 0), _
              C_display_wenti.m_point_no(num, 1), _
                C_display_wenti.m_point_no(num, 2), _
                  0, 0, 0, 1, 1, condition, condition_color, True)
    Call C_display_wenti.set_m_inner_circ(num, tc%, 1)
 End If
     m_Circ(C_display_wenti.m_inner_circ(num, 1)).data(0).from_wenti_no = num
     Call set_element_depend(circle_, C_display_wenti.m_inner_circ(num, 1), _
            point_, C_display_wenti.m_point_no(num, 0), _
             point_, C_display_wenti.m_point_no(num, 1), _
              point_, C_display_wenti.m_point_no(num, 2), True)
           Call set_element_depend(point_, _
            m_Circ(C_display_wenti.m_inner_circ(num, 1)).data(0).data0.center, _
             point_, C_display_wenti.m_point_no(num, 0), _
              point_, C_display_wenti.m_point_no(num, 1), _
               point_, C_display_wenti.m_point_no(num, 2), True)
            m_poi(m_Circ(C_display_wenti.m_inner_circ(num, 1)).data(0).data0.center).data(0).degree = 0
   If C_display_wenti.m_point_no(num, 3) > 0 Then
    temp_record.record_data.data0.condition_data.condition_no = 0 'record0
     tn_% = 0
     Call set_equal_dline(m_Circ(C_display_wenti.m_inner_circ(num, 1)).data(0).data0.center, _
                            C_display_wenti.m_point_no(num, 0), _
                          m_Circ(C_display_wenti.m_inner_circ(num, 1)).data(0).data0.center, _
                            C_display_wenti.m_point_no(num, 1), 0, 0, 0, 0, 0, 0, 0, _
                          temp_record, tn_%, 0, 0, 0, 0, False)
     Call C_display_wenti.set_m_condition_data(num, eline_, tn_%)
     Call set_equal_dline(m_Circ(C_display_wenti.m_inner_circ(num, 1)).data(0).data0.center, _
                            C_display_wenti.m_point_no(num, 0), _
                          m_Circ(C_display_wenti.m_inner_circ(num, 1)).data(0).data0.center, _
                            C_display_wenti.m_point_no(num, 2), 0, 0, 0, 0, 0, 0, 0, _
                          temp_record, 0, 0, 0, 0, 0, False)
     Call set_equal_dline(m_Circ(C_display_wenti.m_inner_circ(num, 1)).data(0).data0.center, _
                            C_display_wenti.m_point_no(num, 2), _
                          m_Circ(C_display_wenti.m_inner_circ(num, 1)).data(0).data0.center, _
                            C_display_wenti.m_point_no(num, 1), 0, 0, 0, 0, 0, 0, 0, _
                          temp_record, 0, 0, 0, 0, 0, False)
   End If
End Sub

Private Sub set_initial_condition9(ByVal num As Integer, temp_record As total_record_type)
'9 直线□□和直线□□交于点□
'depend_no(num) = 1
'Call draw_picture9(num, no_reduce)
Dim i%, j%, k%, l%, no%, bra%
Dim tn(3) As Integer
Dim v(3) As String
Dim tl(3) As Integer
Dim tn_(3) As Integer
Dim v_(1) As String
Dim re  As String
Dim c_data As condition_data_type
Dim lv As V_line_value_data0_type
Dim p4_c As four_point_on_circle_data_type
t_line(0) = C_display_wenti.m_inner_lin(num, 1)
t_line(1) = C_display_wenti.m_inner_lin(num, 2)
Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
                  line_, t_line(0), line_, t_line(1), 0, 0, False)
If regist_data.run_type = 1 Then
  Call set_V_coordinate_system(C_display_wenti.m_point_no(num, 0), _
       C_display_wenti.m_point_no(num, 1))
  Call set_V_coordinate_system(C_display_wenti.m_point_no(num, 2), _
       C_display_wenti.m_point_no(num, 3))
 If is_V_line_value(C_display_wenti.m_point_no(num, 0), _
      C_display_wenti.m_point_no(num, 2), 0, 0, 0, v(0), tn(0), _
       -1000, 0, 0, 0, lv, False) Then
  If is_V_line_value(C_display_wenti.m_point_no(num, 2), _
      C_display_wenti.m_point_no(num, 1), 0, 0, 0, v(1), tn(1), _
       -1000, 0, 0, 0, lv, False) Then
   If is_V_line_value(C_display_wenti.m_point_no(num, 1), _
      C_display_wenti.m_point_no(num, 3), 0, 0, 0, v(2), tn(2), _
       -1000, 0, 0, 0, lv, False) Then
    If is_V_line_value(C_display_wenti.m_point_no(num, 3), _
      C_display_wenti.m_point_no(num, 0), 0, 0, 0, v(3), tn(3), _
       -1000, 0, 0, 0, lv, False) Then
       '******************************************************************
        temp_record.record_data.data0.condition_data.condition_no = 0
        Call add_conditions_to_record(V_line_value_, tn(0), tn(3), 0, _
                   temp_record.record_data.data0.condition_data)
        Call set_item0(V_line_value(tn(3)).data(0).v_line, -10, _
               V_line_value(tn(0)).data(0).v_line, -10, "C", 0, 0, 0, 0, _
                 0, 0, "1", "1", "", cross_time_v_string(v(0), v(3)), v_(0), 0, _
                  temp_record.record_data.data0.condition_data, 0, tn_(0), 0, 0, c_data, True)
        temp_record.record_data.data0.condition_data.condition_no = 0
        Call add_conditions_to_record(V_line_value_, tn(2), tn(1), 0, _
                   temp_record.record_data.data0.condition_data)
        Call set_item0(V_line_value(tn(1)).data(0).v_line, -10, _
               V_line_value(tn(2)).data(0).v_line, -10, "C", 0, 0, 0, 0, _
                 0, 0, "1", "1", "", cross_time_v_string(v(2), v(1)), v_(1), 0, _
                  temp_record.record_data.data0.condition_data, 0, tn_(1), 0, 0, c_data, True)
        temp_record.record_data.data0.condition_data.condition_no = 0
        Call add_conditions_to_record(item0_, tn_(0), tn_(1), 0, _
                   temp_record.record_data.data0.condition_data)
        re = divide_string(v_(0), v_(1), True, False)
        Call set_Drelation(C_display_wenti.m_point_no(num, 0), _
              C_display_wenti.m_point_no(num, 4), _
               C_display_wenti.m_point_no(num, 4), _
                C_display_wenti.m_point_no(num, 1), _
                 0, 0, 0, 0, 0, 0, re, temp_record, 0, 0, 0, 0, 0, False)
        're(0) = divide_string(cross_time_v_string(v(0), v(3)), _
                cross_time_v_string(v(2), v(1)), True, False)
 '***********************************************************************************
        temp_record.record_data.data0.condition_data.condition_no = 0
        Call add_conditions_to_record(V_line_value_, tn(0), tn(1), 0, _
                   temp_record.record_data.data0.condition_data)
        Call set_item0(V_line_value(tn(0)).data(0).v_line, -10, _
               V_line_value(tn(1)).data(0).v_line, -10, "C", 0, 0, 0, 0, _
                 0, 0, "1", "1", "", cross_time_v_string(v(0), v(1)), v_(0), 0, _
                  temp_record.record_data.data0.condition_data, 0, tn_(0), 0, 0, c_data, True)
        temp_record.record_data.data0.condition_data.condition_no = 0
        Call add_conditions_to_record(V_line_value_, tn(2), tn(3), 0, _
                   temp_record.record_data.data0.condition_data)
        Call set_item0(V_line_value(tn(2)).data(0).v_line, -10, _
               V_line_value(tn(3)).data(0).v_line, -10, "C", 0, 0, 0, 0, _
                 0, 0, "1", "1", "", cross_time_v_string(v(2), v(3)), v_(1), 0, _
                  temp_record.record_data.data0.condition_data, 0, tn_(1), 0, 0, c_data, True)
        temp_record.record_data.data0.condition_data.condition_no = 0
        Call add_conditions_to_record(item0_, tn_(0), tn_(1), 0, _
                   temp_record.record_data.data0.condition_data)
        'Call set_Drelation(C_display_wenti.m_point_no(num,0), _
              C_display_wenti.m_point_no(num,4), _
               C_display_wenti.m_point_no(num,4), _
                C_display_wenti.m_point_no(num,1), _
                 0, 0, 0, 0, 0, 0, re(0), temp_record, 0, 0, 0, 0, 0)
        re = divide_string(v_(0), v_(1), True, False)
        Call set_Drelation(C_display_wenti.m_point_no(num, 2), _
              C_display_wenti.m_point_no(num, 4), _
               C_display_wenti.m_point_no(num, 4), _
                C_display_wenti.m_point_no(num, 3), _
                 0, 0, 0, 0, 0, 0, re, temp_record, 0, 0, 0, 0, 0, False)
      End If
      End If
      End If
      End If
      For i% = 2 To m_lin(t_line(0)).data(0).data0.in_point(0)
      If m_lin(t_line(0)).data(0).data0.in_point(i%) <> _
            C_display_wenti.m_point_no(num, 4) Then
      For j% = 1 To i% - 1
      If m_lin(t_line(0)).data(0).data0.in_point(j%) <> _
            C_display_wenti.m_point_no(num, 4) Then
       For k% = 2 To m_lin(t_line(1)).data(0).data0.in_point(0)
        If m_lin(t_line(1)).data(0).data0.in_point(k%) <> _
            C_display_wenti.m_point_no(num, 4) Then
        For l% = 1 To k% - 1
         If m_lin(t_line(1)).data(1).data0.in_point(l%) <> _
            C_display_wenti.m_point_no(num, 4) Then
             If is_four_point_on_circle(m_lin(t_line(0)).data(0).data0.in_point(i%), _
                  m_lin(t_line(0)).data(0).data0.in_point(j%), _
                   m_lin(t_line(1)).data(0).data0.in_point(k%), _
                    m_lin(t_line(1)).data(0).data0.in_point(l%), no%, _
                       p4_c, False) Then
                Call add_conditions_to_record(point4_on_circle_, no%, 0, 0, temp_record.record_data.data0.condition_data)
                Call set_four_point_on_circle_for_vector0(m_lin(t_line(0)).data(0).data0.in_point(i%), _
                   m_lin(t_line(0)).data(0).data0.in_point(j%), _
                    m_lin(t_line(1)).data(0).data0.in_point(k%), _
                     m_lin(t_line(1)).data(0).data0.in_point(l%), temp_record)
             End If
         End If
        Next l%
        End If
       Next k%
       End If
      Next j%
      End If
     Next i%
End If
End Sub

Private Sub set_initial_condition10_16(ByVal num As Integer, temp_record As total_record_type)
'10 过□平行□□的直线交⊙□[down\\(_)]于□
'16 过□垂直□□的直线交⊙□[down\\(_)]于□
'-68 过□垂直□□的直线交⊙□□□于□
'-62 过□平行□□的直线交⊙□□□于□
Dim t_degree As Byte
'depend_no(num) = 1
t_line(0) = C_display_wenti.m_inner_lin(num, 3)
t_line(1) = C_display_wenti.m_inner_lin(num, 4)
Call set_element_depend(line_, t_line(0), _
                         point_, C_display_wenti.m_point_no(num, 1), _
                          line_, t_line(0), 0, 0, False)
Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
                         line_, t_line(0), circle_, _
                          C_display_wenti.m_inner_circ(num, 1), 0, 0, True)
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
  If C_display_wenti.m_no(num) = 10 Or _
           C_display_wenti.m_no(num) = 62 Then
  Call set_dparal(t_line(0), t_line(1), temp_record, 0, 0, True)
  Else
  Call set_dverti(t_line(0), t_line(1), temp_record, 0, 0, True)
  End If
End Sub

Private Sub set_initial_condition12(ByVal num As Integer, temp_record As total_record_type)
'⊙□[down\\(_)]和⊙□[down\\(_)]相切于点□
Dim p(1) As Integer
Dim tl(1) As Integer
Dim tn_%
Dim c_data0 As condition_data_type
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
 tn_% = 0
 Call set_tangent_line(C_display_wenti.m_inner_lin(num, 1), _
                        C_display_wenti.m_inner_poi(num, 1), _
                         C_display_wenti.m_inner_circ(num, 1), _
                          C_display_wenti.m_inner_poi(num, 1), _
                           C_display_wenti.m_inner_circ(num, 2), temp_record, tn_%, 0)
 Call C_display_wenti.set_m_condition_data(num, tangent_line_, tn_%)
 If C_display_wenti.m_inner_lin(num, 1) And _
       C_display_wenti.m_inner_lin(num, 2) > 0 Then
  Call set_dverti(C_display_wenti.m_inner_lin(num, 1), _
                   C_display_wenti.m_inner_lin(num, 2), _
                    temp_record, 0, 0, True)
End If
End Sub

Private Sub set_initial_condition13(ByVal num As Integer, temp_record As total_record_type)
'□是⊙□[down\\(_)]和⊙□[down\\(_)]一个交点
'Call draw_picture13(num, no_reduce)
Dim c1%, c2%
 c1% = C_display_wenti.m_inner_circ(num, 1)
 c2% = C_display_wenti.m_inner_circ(num, 2)
 Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
               circle_, c1%, circle_, c2%, 0, 0, False)
End Sub

Private Sub set_point_from_wenti_no(ByVal num%, ByVal num_no%)
Dim i%, p%
If num_no% < 23 Then
For i% = 0 To 50
 If C_display_wenti.m_point_no(num, i%) > 0 Then
    p% = C_display_wenti.m_point_no(num, i%)
    If C_display_wenti.m_condition(num, i%) >= "A" And _
         C_display_wenti.m_condition(num, i%) <= "Z" Then
          If m_poi(p%).data(0).from_wenti_no = 0 Then
             m_poi(p%).data(0).from_wenti_no = num
              If num_no% > 22 Then
               
              End If
          End If
    End If
 End If
Next i%
End If
End Sub
Private Sub set_initial_data_for_draw_2(ByVal num%)
If m_Circ(C_display_wenti.m_inner_circ(num, 1)).data(0).degree = 0 And _
    m_Circ(C_display_wenti.m_inner_circ(num, 2)).data(0).degree = 0 Then
    m_poi(C_display_wenti.m_inner_poi(num, 1)).data(0).degree_for_reduce = 0
    m_poi(C_display_wenti.m_inner_poi(num, 2)).data(0).degree_for_reduce = 0
ElseIf (m_Circ(C_display_wenti.m_inner_circ(num, 1)).data(0).degree >= 1 And _
    m_Circ(C_display_wenti.m_inner_circ(num, 2)).data(0).degree = 0) Or _
    (m_Circ(C_display_wenti.m_inner_circ(num, 1)).data(0).degree = 0 And _
    m_Circ(C_display_wenti.m_inner_circ(num, 2)).data(0).degree >= 1) Then
    m_poi(C_display_wenti.m_inner_poi(num, 1)).data(0).degree_for_reduce = 1
    m_poi(C_display_wenti.m_inner_poi(num, 2)).data(0).degree_for_reduce = 1
ElseIf m_Circ(C_display_wenti.m_inner_circ(num, 1)).data(0).degree >= 1 And _
    m_Circ(C_display_wenti.m_inner_circ(num, 2)).data(0).degree >= 1 Then
    m_poi(C_display_wenti.m_inner_poi(num, 1)).data(0).degree_for_reduce = 2
    m_poi(C_display_wenti.m_inner_poi(num, 2)).data(0).degree_for_reduce = 2
End If
    If m_poi(C_display_wenti.m_inner_poi(num, 1)).data(0).degree_for_reduce = 0 And _
      m_poi(C_display_wenti.m_inner_poi(num, 2)).data(0).degree_for_reduce = 0 Then
       m_lin(C_display_wenti.m_inner_lin(num, 1)).data(0).degree = 0
    ElseIf (m_poi(C_display_wenti.m_inner_poi(num, 1)).data(0).degree_for_reduce >= 1 And _
             m_poi(C_display_wenti.m_inner_poi(num, 2)).data(0).degree_for_reduce = 0) Or _
             (m_poi(C_display_wenti.m_inner_poi(num, 1)).data(0).degree_for_reduce = 0 And _
               m_poi(C_display_wenti.m_inner_poi(num, 2)).data(0).degree_for_reduce >= 1) Then
       m_lin(C_display_wenti.m_inner_lin(num, 1)).data(0).degree = 1
    ElseIf m_poi(C_display_wenti.m_inner_poi(num, 1)).data(0).degree_for_reduce >= 1 And _
         m_poi(C_display_wenti.m_inner_poi(num, 2)).data(0).degree_for_reduce >= 1 Then
              m_lin(C_display_wenti.m_inner_lin(num, 1)).data(0).degree = 2
    End If
End Sub
Private Sub set_initial_data_for_draw_1(ByVal num%)
'-1 □□＝□□
Dim tp(3) As Integer
Dim i%, c_n%
Dim p_coord(2) As POINTAPI
Dim r(1) As Single
If C_display_wenti.m_point_no(num, 30) > 0 Then
   Exit Sub
End If
For i% = 0 To 3
 tp(i%) = C_display_wenti.m_point_no(num, i%)
Next i%
If m_poi(tp(0)).data(0).degree_for_reduce = 0 And _
     m_poi(tp(1)).data(0).degree_for_reduce = 0 And _
       m_poi(tp(2)).data(0).degree_for_reduce = 0 And _
         m_poi(tp(3)).data(0).degree_for_reduce = 0 Then
If m_poi(tp(3)).data(0).degree = 2 Then
ElseIf m_poi(tp(2)).data(0).degree = 2 Then
    Call exchange_two_integer(tp(2), tp(3))
ElseIf m_poi(tp(1)).data(0).degree = 2 Then
     Call exchange_two_integer(tp(0), tp(2))
     Call exchange_two_integer(tp(1), tp(3))
ElseIf m_poi(tp(0)).data(0).degree = 2 Then
   Call exchange_two_integer(tp(0), tp(3))
   Call exchange_two_integer(tp(1), tp(3))
End If
Else
 If m_poi(tp(0)).data(0).degree_for_reduce > m_poi(tp(1)).data(0).degree_for_reduce Then
   Call exchange_two_integer(tp(0), tp(1))
 End If
 If m_poi(tp(2)).data(0).degree_for_reduce > m_poi(tp(3)).data(0).degree_for_reduce Then
   Call exchange_two_integer(tp(2), tp(3))
 End If
 If m_poi(tp(1)).data(0).degree_for_reduce > m_poi(tp(3)).data(0).degree_for_reduce Then
   Call exchange_two_integer(tp(0), tp(2))
   Call exchange_two_integer(tp(1), tp(3))
 End If
End If
'p3%是自由度最大的点
If tp(0) = tp(3) Then
   Call exchange_two_integer(tp(0), tp(1))
   Call C_display_wenti.set_m_point_no(num, 3, 48, False)
ElseIf tp(1) = tp(3) Then
   Call C_display_wenti.set_m_point_no(num, 3, 48, False)
End If
'******************************
For i% = 0 To 3
 Call C_display_wenti.set_m_point_no(num, tp(i%), 30 + i%, False)
Next i%
If C_display_wenti.m_point_no(num, 48) = 3 Then
   If m_poi(tp(3)).data(0).parent.element(0).ty = line_ Then
    If inter_point_verti_mid_line_line(m_poi(tp(0)).data(0).data0.coordinate, _
             m_poi(tp(2)).data(0).data0.coordinate, _
                m_lin(m_poi(tp(3)).data(0).parent.element(0).no).data(0).data0, _
                 p_coord(0)) Then
      Call set_point_coordinate(tp(3), p_coord(0), True)   '
    End If
   ElseIf m_poi(tp(3)).data(0).parent.element(0).ty = circle_ Then
    If inter_point_verti_mid_line_circle(m_poi(tp(0)).data(0).data0.coordinate, _
             m_poi(tp(2)).data(0).data0.coordinate, _
                m_Circ(m_poi(tp(3)).data(0).parent.element(0).no).data(0).data0, _
                 p_coord(0), p_coord(1)) Then
          p_coord(2) = minus_POINTAPI(m_poi(tp(3)).data(0).data0.coordinate, _
                 p_coord(0))
          r(0) = time_POINTAPI(p_coord(2), p_coord(2))
          p_coord(2) = minus_POINTAPI(m_poi(tp(3)).data(0).data0.coordinate, _
                 p_coord(1))
          r(1) = time_POINTAPI(p_coord(2), p_coord(2))
             Call C_display_wenti.set_m_point_no(num, 5, 48, False)
          If r(0) < r(1) Then
            Call set_point_coordinate(tp(3), p_coord(0), True)  '
            Call C_display_wenti.set_m_point_no(num, 1, 47, False)
          Else
           Call set_point_coordinate(tp(3), p_coord(1), True)  '
            Call C_display_wenti.set_m_point_no(num, 2, 47, False)
          End If
    End If
   Else
            Call C_display_wenti.set_m_point_no(num, 4, 48, False)
    Call inter_point_verti_mid_line_point(m_poi(tp(0)).data(0).data0.coordinate, _
             m_poi(tp(2)).data(0).data0.coordinate, _
                 m_poi(tp(3)).data(0).data0.coordinate, _
                 p_coord(0))
      Call set_point_coordinate(tp(3), p_coord(0), True)  '
   End If
Else
m_input_circle_data0.data0.center = tp(2)
 m_input_circle_data0.data0.radii = abs_POINTAPI(minus_POINTAPI( _
   m_poi(tp(0)).data(0).data0.coordinate, m_poi(tp(1)).data(0).data0.coordinate))
c_n% = m_circle_number(1, tp(2), pointapi0, tp(3), 0, 0, 0, tp(0), tp(1), 1, 0, condition, condition_color, True)
If m_poi(tp(3)).data(0).parent.element(0).no = 0 Then
   m_poi(tp(3)).data(0).parent.element(0).no = c_n%
    m_poi(tp(3)).data(0).parent.element(0).ty = circle_
    'If m_poi(tp(3)).data(0).parent.element(0).ty <> line_ And _
        m_poi(tp(3)).data(0).parent.element(1).ty <> line_ Then
     m_poi(tp(3)).data(0).degree_for_reduce = 1
       r(0) = abs_POINTAPI(minus_POINTAPI( _
         m_poi(tp(2)).data(0).data0.coordinate, m_poi(tp(3)).data(0).data0.coordinate))
       p_coord(0) = minus_POINTAPI(m_poi(tp(3)).data(0).data0.coordinate, _
                         m_poi(tp(2)).data(0).data0.coordinate)
       p_coord(0) = time_POINTAPI_by_number(p_coord(0), m_Circ(c_n%).data(0).data0.radii / r(0))
       p_coord(0) = add_POINTAPI(p_coord(0), m_poi(tp(2)).data(0).data0.coordinate)
       'p_coord(0).X = m_poi(tp(2)).data(0).data0.coordinate.X + _
           (m_poi(tp(3)).data(0).data0.coordinate.X - m_poi(tp(2)).data(0).data0.coordinate.X) * _
             m_Circ(C_n%).data(0).data0.radii / r(0)
       'p_coord(0).Y = m_poi(tp(2)).data(0).data0.coordinate.Y + _
           (m_poi(tp(3)).data(0).data0.coordinate.Y - m_poi(tp(2)).data(0).data0.coordinate.Y) * _
             m_Circ(C_n%).data(0).data0.radii / r(0)
       Call set_point_coordinate(tp(3), p_coord(0), True)  '
    ElseIf m_poi(tp(3)).data(0).parent.element(0).ty = line_ And _
               m_poi(tp(3)).data(0).parent.element(1).ty = circle_ Then
     m_poi(tp(3)).data(0).degree_for_reduce = 0
          Call C_display_wenti.set_m_point_no(num, 1, 48, False)  '
      Call inter_point_line_circle1( _
          m_poi(m_lin(m_poi(tp(3)).data(0).parent.element(0).no).data(0).data0.poi(0)).data(0).data0.coordinate, _
           paral_, m_lin(m_poi(tp(3)).data(0).parent.element(0).no).data(0), _
            m_Circ(m_poi(tp(3)).data(0).parent.element(1).no).data(0).data0, _
              p_coord(0), 0, p_coord(1), 0)
       r(0) = (m_poi(tp(3)).data(0).data0.coordinate.X - p_coord(0).X) ^ 2 + _
                       (m_poi(2).data(0).data0.coordinate.Y - p_coord(0).Y) ^ 2
       r(1) = (m_poi(tp(3)).data(0).data0.coordinate.X - p_coord(1).X) ^ 2 + _
                       (m_poi(2).data(0).data0.coordinate.Y - p_coord(1).Y) ^ 2
        If r(0) < r(1) Then
          Call C_display_wenti.set_m_point_no(num, 1, 47, False)
           Call set_point_coordinate(tp(3), p_coord(0), True)   '
        Else
           Call C_display_wenti.set_m_point_no(num, 2, 47, False)
             Call set_point_coordinate(tp(3), p_coord(1), True)  '
        End If
    'End If
Else
   m_poi(tp(3)).data(0).parent.element(1).no = c_n%
    m_poi(tp(3)).data(0).parent.element(1).ty = circle_
      If m_poi(tp(3)).data(0).parent.element(0).ty <> line_ Then
    m_poi(tp(3)).data(0).degree_for_reduce = 0
          Call C_display_wenti.set_m_point_no(num, 2, 48, False)  '
     Call inter_point_circle_circle_(m_Circ(m_poi(tp(3)).data(0).parent.element(0).no).data(0).data0, _
            m_Circ(m_poi(tp(3)).data(0).parent.element(1).no).data(0).data0, _
               p_coord(0), 0, p_coord(1), 0, 0, 0, False)
      r(0) = (m_poi(tp(3)).data(0).data0.coordinate.X - p_coord(0).X) ^ 2 + _
                       (m_poi(2).data(0).data0.coordinate.Y - p_coord(0).Y) ^ 2
      r(1) = (m_poi(tp(3)).data(0).data0.coordinate.X - p_coord(1).X) ^ 2 + _
                       (m_poi(2).data(0).data0.coordinate.Y - p_coord(1).Y) ^ 2
      'm_point_data0 = m_poi(tp(3)).data(0).data0
      If r(0) < r(1) Then
      Call C_display_wenti.set_m_point_no(num, 1, 47, False)
       Call set_point_coordinate(tp(3), p_coord(0), True)  '
      Else
      Call C_display_wenti.set_m_point_no(num, 2, 47, False)
       Call set_point_coordinate(tp(3), p_coord(1), True)  '
      End If
      Else
        error_of_wenti = 1
      End If
End If
If m_poi(tp(3)).data(0).parent.element(0).ty = circle_ Then
Call C_display_wenti.set_m_point_no(num, m_poi(tp(3)).data(0).parent.element(0).no, 42, False)
End If
If m_poi(tp(3)).data(0).parent.element(1).ty = circle_ Then
Call C_display_wenti.set_m_point_no(num, m_poi(tp(3)).data(0).parent.element(1).no, 43, False)
ElseIf m_poi(tp(3)).data(0).parent.element(0).ty = line_ Then
Call C_display_wenti.set_m_point_no(num, m_poi(tp(3)).data(0).parent.element(0).no, 40, False)
End If
End If
Call C_display_wenti.set_m_point_no(num, tp(3), 44, False)
End Sub

Public Sub set_initial_condition11_63(ByVal num As Integer, temp_record As total_record_type)
'11 □是直线□□与⊙□[down\\(_)]的一个交点
'-63 □是直线□□与⊙□□□的一个交点
Dim t_vl(2) As Integer
Dim c%, l%, tp%
Dim tn(2) As Integer
Dim para(2) As String
Dim dir(1) As String
Dim c_data As condition_data_type
  c% = C_display_wenti.m_inner_circ(num, 1)
  l% = C_display_wenti.m_inner_lin(num, 1)
  tp% = C_display_wenti.m_inner_poi(num, 1)
  Call set_element_depend(point_, C_display_wenti.m_point_no(num, 0), _
       line_, l%, circle_, c%, 0, 0, False)
'*******************************************
If regist_data.run_type = 1 Then
            If C_display_wenti.m_point_no(num, 1) <> _
                     tp% Then
               Call set_V_coordinate_system(m_Circ(c%).data(0).data0.center, _
                  C_display_wenti.m_point_no(num, 1)) '圆心和线上的点连线
               If is_point_in_circle(c%, 0, C_display_wenti.m_point_no(num, 1), 0, 0) Then
                      If m_Circ(c%).data(0).parent.element(1).no = 0 Then
                          m_Circ(c%).data(0).parent.element(2).no = _
                             m_Circ(c%).data(0).data0.center
                          m_Circ(c%).data(0).parent.element(1).no = _
                              C_display_wenti.m_point_no(num, 1)
                      End If
               End If
            ElseIf C_display_wenti.m_point_no(num, 2) <> tp% Then
               Call set_V_coordinate_system(m_Circ(C_display_wenti.m_point_no(num, 42)).data(0).data0.center, _
                  C_display_wenti.m_point_no(num, 2))
               If is_point_in_circle(C_display_wenti.m_point_no(num, 42), 0, _
                      C_display_wenti.m_point_no(num, 2), 0, 0) Then
                      If m_Circ(c%).data(0).parent.element(1).no = 0 Then
                          m_Circ(c%).data(0).parent.element(2).no = _
                             m_Circ(c%).data(0).data0.center
                          m_Circ(c%).data(0).parent.element(1).no = _
                              C_display_wenti.m_point_no(num, 2)
                      End If
               End If
            End If
     t_vl(0) = vector_number(C_display_wenti.m_point_no(num, 1), _
                              C_display_wenti.m_point_no(num, 2), 0) '连线
     t_vl(1) = vector_number(C_display_wenti.m_point_no(num, 1), _
                             m_Circ(c%).data(0).data0.center, dir(0))
     t_vl(2) = vector_number(C_display_wenti.m_point_no(num, 2), _
                              m_Circ(c%).data(0).data0.center, dir(1))
     Call set_item0(t_vl(0), -10, t_vl(1), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), 0, _
           c_data, 0, tn(0), 0, 0, c_data, False)
     Call set_item0(t_vl(0), -10, t_vl(2), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), 0, _
           c_data, 0, tn(1), 0, 0, c_data, False)
      para(0) = time_string(dir(0), para(0), True, False)
      para(1) = time_string(dir(1), para(1), True, False)
     Call set_general_string(tn(0), tn(1), 0, 0, para(0), para(1), "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
          t_vl(0) = vector_number(tp%, _
              m_Circ(c%).data(0).data0.center, dir(0))
     t_vl(1) = vector_number(m_Circ(c%).data(0).parent.element(1).no, _
              m_Circ(c%).data(0).data0.center, dir(0))
     Call set_item0(t_vl(0), -10, t_vl(0), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), 0, _
           c_data, 0, tn(0), 0, 0, c_data, False)
     Call set_item0(t_vl(1), -10, t_vl(1), -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), 0, _
           c_data, 0, tn(1), 0, 0, c_data, False)
     Call set_general_string(tn(0), tn(1), 0, 0, para(0), _
                   time_string("-1", para(1), True, False), "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
End If
End Sub
Private Function set_V_coordinate_system(ByVal p1%, ByVal p2%) As Boolean
Dim Vul As String
Dim temp_record As total_record_type
temp_record.record_data.data0.condition_data.condition(8).ty = new_point_
If (m_poi(p1%).data(0).degree_for_reduce <= 1 And m_poi(p1%).data(0).degree >= 1) And _
    (m_poi(p2%).data(0).degree_for_reduce <= 1 And m_poi(p2%).data(0).degree >= 1) Then
If last_conditions.last_cond(1).init_v_line_no = 0 Then
        Vul = "U"
        set_V_coordinate_system = True
ElseIf last_conditions.last_cond(1).init_v_line_no = 1 Then
       If is_dparal( _
           Dtwo_point_line(V_line_value(v_coordinate_system_no(0)).data(0).v_line).data(0).line_no, _
             line_number0(p1%, p2%, 0, 0), 0, -1000, 0, 0, 0, 0) Then
              Exit Function
       Else
        Vul = "V"
         set_V_coordinate_system = True
       End If
Else
        Exit Function
End If
        v_coordinate_system_no(last_conditions.last_cond(1).init_v_line_no) = 0
        Call set_V_line_value(p1%, p2%, 0, 0, 0, Vul, temp_record, _
               v_coordinate_system_no(last_conditions.last_cond(1).init_v_line_no), True)
        last_conditions.last_cond(1).init_v_line_no = last_conditions.last_cond(1).init_v_line_no + 1
End If
End Function
Public Sub set_initial_condition2_3(ByVal num As Integer, temp_record As total_record_type)
'2 □□∥□□
'3 □□⊥□□
Dim i%, tn_%, tn%
Dim para$
Dim tp(3) As Integer
Dim t_line(1) As Integer
Dim c_data As condition_data_type
For i% = 0 To 3
tp(i%) = C_display_wenti.m_point_no(num, i%)
Next i%
If C_display_wenti.m_inner_poi(num, 1) = 0 Then
   Call set_wenti_cond2_3(tp(2), tp(0), tp(1), tp(3), _
          C_display_wenti.m_no(num), num)
End If
t_line(0) = C_display_wenti.m_inner_lin(num, 3)
t_line(1) = C_display_wenti.m_inner_lin(num, 4)
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
If C_display_wenti.m_no(num) = 2 Then
 tn_% = 0
 Call set_dparal(t_line(0), t_line(1), temp_record, tn_%, 0, True)
 Call C_display_wenti.set_m_condition_data(num, paral_, tn_%)
Else
 tn_% = 0
 Call set_dverti(t_line(0), t_line(1), temp_record, tn_%, 0, True)
 Call C_display_wenti.set_m_condition_data(num, verti_, tn_%)
End If
If C_display_wenti.m_inner_point_type(num) = 0 Then
   Call set_element_depend(line_, t_line(0), _
         point_, C_display_wenti.m_inner_poi(num, 2), _
          line_, t_line(1), 0, 0, False)
   Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
         line_, t_line(0), 0, 0, 0, 0, False)
Else
   If C_display_wenti.m_inner_lin(num, 1) > 0 Then
   Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
         line_, t_line(0), line_, C_display_wenti.m_inner_lin(num, 1), 0, 0, True)
   ElseIf C_display_wenti.m_inner_circ(num, 1) > 0 Then
   Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), _
         line_, t_line(0), line_, C_display_wenti.m_inner_circ(num, 1), 0, 0, True)
   End If
End If
 If regist_data.run_type = 1 Then
   Call set_V_coordinate_system(C_display_wenti.m_point_no(num, 0), _
                                 C_display_wenti.m_point_no(num, 1))
   Call set_V_coordinate_system(C_display_wenti.m_point_no(num, 2), _
                                 C_display_wenti.m_point_no(num, 3))
  If C_display_wenti.m_no(num) = 2 Then
  Else
   t_line(0) = vector_number(C_display_wenti.m_point_no(num, 0), _
              C_display_wenti.m_point_no(num, 1), "")
   t_line(1) = vector_number(C_display_wenti.m_point_no(num, 2), _
              C_display_wenti.m_point_no(num, 3), "")
      temp_record.record_data.data0.condition_data.condition_no = 0
      Call add_conditions_to_record(verti_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
      Call set_item0(t_line(0), -10, t_line(1), -10, "*", 0, 0, 0, 0, 0, _
           "1", "1", "1", "1", "", para$, 0, temp_record.record_data.data0.condition_data, _
             0, tn%, 0, 0, temp_record.record_data.data0.condition_data, False)
      If tn% > 0 Then
       Call set_general_string(tn%, 0, 0, 0, para$, "0", "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
      Else
       Call set_relation_string(para$, 0, temp_record)
      End If
  End If
 End If
End Sub
Public Sub set_initial_condition14(ByVal num As Integer, temp_record As total_record_type)
'14过□作直线□□的垂线垂足为□
Dim i%, j%, tn%
Dim para$
temp_record.record_data.data0.condition_data.condition_no = 0 'record0
 i% = C_display_wenti.m_inner_lin(num, 1)
 j% = C_display_wenti.m_inner_lin(num, 2)
 Call set_element_depend(line_, j%, line_, _
                i%, point_, C_display_wenti.m_point_no(num, 0), 0, 0, True)
 Call set_element_depend(point_, C_display_wenti.m_inner_poi(num, 1), line_, _
                j%, line_, i%, 0, 0, False)
 Call set_dverti(i%, j%, temp_record, tn%, 0, True)
Call set_line_degree(i%)
If m_poi(C_display_wenti.m_point_no(num, 1)).data(0).degree = 2 Or _
      m_poi(C_display_wenti.m_point_no(num, 1)).data(0).degree = 2 Then
        m_poi(C_display_wenti.m_point_no(num, 1)).data(0).degree_for_reduce = 0
Else
      m_poi(C_display_wenti.m_point_no(num, 1)).data(0).degree_for_reduce = 1
End If
If m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree = 2 Or _
      m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree = 2 Then
        m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce = 0
Else
      m_poi(C_display_wenti.m_point_no(num, 2)).data(0).degree_for_reduce = 1
End If
If m_poi(C_display_wenti.m_point_no(num, 0)).data(0).degree = 2 Or _
      m_poi(C_display_wenti.m_point_no(num, 0)).data(0).degree = 2 Then
        m_poi(C_display_wenti.m_point_no(num, 0)).data(0).degree_for_reduce = 0
Else
      m_poi(C_display_wenti.m_point_no(num, 0)).data(0).degree_for_reduce = 1
End If
m_poi(C_display_wenti.m_point_no(num, 3)).data(0).degree = 0
m_poi(C_display_wenti.m_point_no(num, 0)).data(0).degree_for_reduce = 0
If regist_data.run_type = 1 Then
   i% = vector_number(C_display_wenti.m_point_no(num, 1), _
         C_display_wenti.m_point_no(num, 2), "")
   j% = vector_number(C_display_wenti.m_point_no(num, 0), _
         C_display_wenti.m_point_no(num, 3), "")
   temp_record.record_data.data0.condition_data.condition_no = 0
   Call add_conditions_to_record(verti_, tn%, 0, 0, temp_record.record_data.data0.condition_data)
   tn% = 0
   Call set_item0(i%, -10, j%, -10, "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", _
           para$, 0, temp_record.record_data.data0.condition_data, 0, tn%, 0, 0, temp_record.record_data.data0.condition_data, False)
   If tn% = 0 Then
   Call set_relation_string(para$, 0, temp_record)
   Else
   Call set_general_string(tn%, 0, 0, 0, para$, "0", "0", "0", "0", 0, 0, 0, _
          temp_record, 0, 0)
   End If
End If
End Sub

Public Function set_four_point_on_circle_for_vector0(ByVal p1%, ByVal p2%, ByVal p3%, _
                                                    ByVal p4%, re As total_record_type) As Byte
Dim p%
p% = is_line_line_intersect(line_number0(p1%, p2%, 0, 0), line_number0(p3%, p4%, 0, 0), 0, 0, False)
 If p% > 0 Then
  set_four_point_on_circle_for_vector0 = set_four_point_on_circle_for_vector(p1%, p2%, p3%, p4%, p%, re)
 Else
  p% = is_line_line_intersect(line_number0(p1%, p3%, 0, 0), line_number0(p2%, p4%, 0, 0), 0, 0, False)
   If p% > 0 Then
    set_four_point_on_circle_for_vector0 = set_four_point_on_circle_for_vector(p1%, p3%, p2%, p4%, p%, re)
   Else
    p% = is_line_line_intersect(line_number0(p1%, p4%, 0, 0), line_number0(p3%, p2%, 0, 0), 0, 0, False)
     If p% > 0 Then
      set_four_point_on_circle_for_vector0 = set_four_point_on_circle_for_vector(p1%, p4%, p2%, p3%, p%, re)
     End If
    End If
 End If
End Function
Public Function set_four_point_on_circle_for_vector(ByVal p1%, ByVal p2%, ByVal p3%, _
                                                    ByVal p4%, ByVal inter_set_p%, _
                                                           re As total_record_type) As Byte
Dim tl(3) As Integer
Dim tn(3) As Integer
Dim it(1) As Integer
Dim dir(3) As String
Dim c_data As condition_data_type
Dim para(1) As String
Dim temp_record As total_record_type
Call set_V_coordinate_system(p1%, p2%)
Call set_V_coordinate_system(p3%, p4%)
temp_record = re
tl(0) = vector_number(p1%, inter_set_p%, dir(0))
tl(1) = vector_number(p2%, inter_set_p%, dir(1))
tl(2) = vector_number(p3%, inter_set_p%, dir(2))
tl(3) = vector_number(p4%, inter_set_p%, dir(3))
Call set_item0(tl(0), -10, tl(1), -10, "*", 0, 0, 0, 0, 0, 0, _
                "1", "1", "1", "", para(0), 0, c_data, 0, tn(0), _
                    0, 0, temp_record.record_data.data0.condition_data, False)
Call set_item0(tl(2), -10, tl(3), -10, "*", 0, 0, 0, 0, 0, 0, _
                "1", "1", "1", "", para(1), 0, c_data, 0, tn(1), _
                    0, 0, temp_record.record_data.data0.condition_data, False)
para(0) = time_string(dir(0), para(0), False, False)
para(0) = time_string(dir(1), para(0), True, False)
para(1) = time_string(dir(2), para(1), False, False)
para(1) = time_string(dir(3), para(1), False, False)
para(1) = time_string("-1", para(1), True, False)
set_four_point_on_circle_for_vector = set_general_string(tn(0), tn(1), 0, 0, _
     para(0), para(1), "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
     If set_four_point_on_circle_for_vector > 0 Then
        Exit Function
     End If
End Function
Public Sub add_new_point_on_circle_for_vector(ByVal c%, ByVal new_point%, _
                                                      re As total_record_type)
Dim i%, j%, k%, n%
For i% = 1 To m_Circ(c%).data(0).data0.in_point(0)
   If m_Circ(c%).data(0).data0.in_point(i%) = new_point Then
      n% = i%
       GoTo add_new_point_on_circle_for_vector_mark0
   End If
Next i%
add_new_point_on_circle_for_vector_mark0:
If n% < 4 Then
   Exit Sub
End If
For i% = 1 To n% - 1
 For j% = 1 To n% - 1
  For k% = 1 To n% - 1
   Call set_four_point_on_circle_for_vector0(m_Circ(c%).data(0).data0.in_point(i%), _
       m_Circ(c%).data(0).data0.in_point(j%), m_Circ(c%).data(0).data0.in_point(k%), _
        new_point%, re)
  Next k%
 Next j%
Next i%
End Sub
Public Sub set_line_degree(ByVal l%)
If m_lin(l%).data(0).parent.element(1).ty = point_ Then
   If (m_poi(m_lin(l%).data(0).parent.element(0).no).data(0).degree = 0 And _
        m_poi(m_lin(l%).data(0).parent.element(1).no).data(0).degree = 0) Or _
      (m_poi(m_lin(l%).data(0).parent.element(0).no).data(0).degree = 2 And _
       m_poi(m_lin(l%).data(0).parent.element(1).no).data(0).degree = 2) Then
        m_lin(l%).data(0).degree = 0
   Else
        m_lin(l%).data(0).degree = 1
   End If
Else
    If (m_poi(m_lin(l%).data(0).parent.element(0).no).data(0).degree = 0 Or _
        m_poi(m_lin(l%).data(0).parent.element(0).no).data(0).degree = 2) And _
         m_lin(m_lin(l%).data(0).parent.element(1).no).data(0).degree = 0 Then
          m_lin(l%).data(0).degree = 0
    Else
          m_lin(l%).data(0).degree = 1
    End If

End If
End Sub
Public Sub set_point_degree(ByVal p%)
If m_poi(p%).data(0).parent.element(0).no = 0 Then
    m_poi(p%).data(0).degree = 2
ElseIf m_poi(p%).data(0).parent.element(1).no = 0 Then
    m_poi(p%).data(0).degree = 1
Else
    m_poi(p%).data(0).degree = 0
End If
End Sub
Public Sub set_circle_degree(ByVal c%)
If m_Circ(c%).data(0).data0.center > m_Circ(c%).data(0).data0.in_point(3) And _
     m_Circ(c%).data(0).data0.in_point(0) >= 3 Then
    If (m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).degree = 0 Or _
        m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).degree = 2) And _
       (m_poi(m_Circ(c%).data(0).data0.in_point(2)).data(0).degree = 0 Or _
         m_poi(m_Circ(c%).data(0).data0.in_point(2)).data(0).degree = 2) And _
       (m_poi(m_Circ(c%).data(0).data0.in_point(3)).data(0).degree = 0 Or _
         m_poi(m_Circ(c%).data(0).data0.in_point(3)).data(0).degree = 2) Then
          m_Circ(c%).data(0).degree = 0
    Else
         m_Circ(c%).data(0).degree = 1
    End If
Else
  If m_poi(m_Circ(c%).data(0).data0.center).data(0).degree = 0 Or _
        m_poi(m_Circ(c%).data(0).data0.center).data(0).degree = 2 Then
    If m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).degree = 0 Or _
        m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).degree = 2 Then
         m_Circ(c%).data(0).degree = 0
    Else
         m_Circ(c%).data(0).degree = 1
    End If
  Else
    'If m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).degree = 0 Or _
        m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).degree = 2 Then
         m_Circ(c%).data(0).degree = 1
    'Else
    '     m_Circ(c%).data(0).degree = 2
    'End If
  End If
End If
End Sub
Public Sub set_data_inform0(ByVal w_n%, ByVal inform$, ByVal in_p$) '
Dim id$
Dim data_ty As Byte
Dim data_no As Integer
Dim data_p%, p1%, p2%
If inform$ = "" Then
   Exit Sub
Else
 id$ = Mid$(inform$, 1, 1)
 data_no = val(Mid$(inform$, 2, 1))
 If id$ = "(" Then
    data_ty = point_
    data_no = C_display_wenti.m_inner_poi(w_n%, data_no)
 ElseIf id$ = "[" Then
    data_ty = line_
    data_no = C_display_wenti.m_inner_lin(w_n%, data_no)
 ElseIf id$ = "{" Then
    data_ty = circle_
    data_no = C_display_wenti.m_inner_circ(w_n%, data_no)
 Else
    Exit Sub
 End If
  inform$ = Mid$(inform$, 4, Len(inform$) - 3)
  in_p$ = ""
  Do While inform <> ""
  id$ = Mid$(inform$, 1, 1)
   If id$ = "[" Then
      data_p% = val(Mid$(inform$, 2, 1))
        inform$ = Mid$(inform$, 4, Len(inform$) - 3)
         in_p$ = in_p$ + C_display_wenti.m_condition(w_n%, data_p%)
   ElseIf id$ = "!" Then
        data_p% = val(Mid$(inform$, 2, 1))
        '????Call read_number_from_wenti(w_n%, data_p%, 0, id$)
                in_p$ = in_p$ + id$
        inform$ = Mid$(inform$, 4, Len(inform$) - 3)
   Else
        in_p$ = in_p$ + id$
        inform$ = Mid$(inform$, 2, Len(inform$) - 1)
   End If
  Loop
  If in_p$ <> "" Then
     If data_ty = point_ Then
        m_poi(data_no%).data(0).inform = in_p$
     ElseIf data_ty = line_ Then
        m_lin(data_no%).data(0).inform = in_p$
     ElseIf data_ty = circle_ Then
        m_Circ(data_no%).data(0).inform = in_p$
     End If
  End If
End If
End Sub

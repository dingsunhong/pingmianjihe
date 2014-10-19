Attribute VB_Name = "information"
Option Explicit
Public Sub set_inform()
If last_conditions.last_cond(1).angle_relation_no > 0 Then
MDIForm1.angle_inform.Enabled = True
MDIForm1.angle_relation.Enabled = True
End If
If last_conditions.last_cond(1).area_of_circle_no > 0 Then
MDIForm1.inform_circle.Enabled = True
MDIForm1.area_of_circle.Enabled = True
End If
If last_conditions.last_cond(1).area_of_element_no > 0 Then
MDIForm1.infrom_polygon.Enabled = True
MDIForm1.area_of_polygon.Enabled = True
End If
If last_conditions.last_cond(1).eangle_no > 0 Then
MDIForm1.angle_inform.Enabled = True
MDIForm1.eangle.Enabled = True
End If
If last_conditions.last_cond(1).eline_no > 0 Or last_conditions.last_cond(1).mid_point_no > 0 Then
MDIForm1.inform_segment.Enabled = True
MDIForm1.eline.Enabled = True
End If
If last_conditions.last_cond(1).four_point_on_circle_no > 0 Then
MDIForm1.inform_circle.Enabled = True
MDIForm1.four_point_on_circle.Enabled = True
End If
If last_conditions.last_cond(1).line_value_no > 0 Then
MDIForm1.inform_segment.Enabled = True
MDIForm1.length_of_segment.Enabled = True
End If
If last_conditions.last_cond(1).paral_no > 0 Then
MDIForm1.inform_line.Enabled = True
MDIForm1.paral.Enabled = True
End If
If last_conditions.last_cond(1).dpoint_pair_no > 0 Then
MDIForm1.inform_segment.Enabled = True
MDIForm1.re_line.Enabled = True
End If
If last_conditions.last_cond(1).relation_no Then
MDIForm1.inform_segment.Enabled = True
MDIForm1.relation.Enabled = True
End If
If last_conditions.last_cond(1).angle_value_no > 0 Then
MDIForm1.angle_inform.Enabled = True
MDIForm1.yizhiA.Enabled = True
End If
If last_conditions.last_cond(1).angle_value_90_no > 0 Then
MDIForm1.angle_inform.Enabled = True
MDIForm1.right_angle.Enabled = True
End If
If last_conditions.last_cond(1).similar_triangle_no > 0 Then
MDIForm1.inform_triangle.Enabled = True
MDIForm1.similar_triangle.Enabled = True
End If
If last_conditions.last_cond(1).tixing_no > 0 Or _
    last_conditions.last_cond(1).parallelogram_no > 0 Or _
     last_conditions.last_cond(1).rhombus_no > 0 Or _
      last_conditions.last_cond(1).long_squre_no > 0 Or _
       last_conditions.last_cond(1).epolygon_no Then
MDIForm1.infrom_polygon.Enabled = True
MDIForm1.sp_four_sides.Enabled = True
End If
If last_conditions.last_cond(1).three_angle_value_sum_no > 0 Then
MDIForm1.angle_inform.Enabled = True
MDIForm1.three_angle.Enabled = True
End If
If last_conditions.last_cond(1).three_point_on_line_no > 0 Then
MDIForm1.inform_line.Enabled = True
MDIForm1.three_point_on_line.Enabled = True
End If
If last_conditions.last_cond(1).total_equal_triangle_no > 0 Then
MDIForm1.inform_triangle.Enabled = True
MDIForm1.total_equal_triangle.Enabled = True
End If
If last_conditions.last_cond(1).two_angle_value_sum_no > 0 Then
MDIForm1.angle_inform.Enabled = True
MDIForm1.two_angle.Enabled = True
End If
If last_conditions.last_cond(1).two_angle_value_180_no > 0 Then
MDIForm1.angle_inform.Enabled = True
MDIForm1.sum_two_angle_pi.Enabled = True
End If
If last_conditions.last_cond(1).two_angle_value_90_no > 0 Then
MDIForm1.angle_inform.Enabled = True
MDIForm1.sum_two_angle_right.Enabled = True
End If
If last_conditions.last_cond(1).two_line_value_no > 0 Then
MDIForm1.inform_segment.Enabled = True
MDIForm1.two_line_value.Enabled = True
End If
If last_conditions.last_cond(1).verti_no > 0 Then
MDIForm1.inform_line.Enabled = True
MDIForm1.verti.Enabled = True
End If
End Sub
Public Sub set_circle_inform(ByVal c%, inf$)
Dim ts As String
If m_Circ(c%).data(0).circle_type = 1 Then
 ts = "¡Ñ" + m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.name
Else
 ts = LoadResString_(1565, "\\1\\" + m_poi(m_Circ(c%).data(0).data0.in_point(1)).data(0).data0.name + _
                          "\\2\\" + m_poi(m_Circ(c%).data(0).data0.in_point(2)).data(0).data0.name + _
                          "\\3\\" + m_poi(m_Circ(c%).data(0).data0.in_point(3)).data(0).data0.name)
End If

If m_Circ(c%).data(0).inform = "" Then
   If inf$ <> "" Then
    m_Circ(c%).data(0).inform = inf$
   ElseIf m_Circ(c%).data(0).circle_type = 2 Or m_Circ(c%).data(0).circle_type = 3 Then
          Call set_point_inform(m_Circ(c%).data(0).data0.center, LoadResString_(1575, "\\1\\" & ts))
   End If
End If
   If m_Circ(c%).data(0).circle_type = 1 Then
       Call set_point_inform(m_Circ(c%).data(0).data0.in_point(1), LoadResString_(1580, "\\1\\" & ts))
       Call set_point_inform(m_Circ(c%).data(0).data0.in_point(2), LoadResString_(1580, "\\1\\" & ts))
       Call set_point_inform(m_Circ(c%).data(0).data0.in_point(3), LoadResString_(1580, "\\1\\" & ts))
   Else
       Call set_point_inform(m_Circ(c%).data(0).data0.in_point(1), LoadResString_(1580, "\\1\\" & ts))
   End If

End Sub
Public Sub set_line_inform(ByVal l%, inf$)
If m_lin(l%).data(0).inform = "" Then
   m_lin(l%).data(0).inform = inf$
End If
End Sub
Public Sub set_information_list(data_type As Byte)
Dim inform_caption As String
Dim i%, n%, k%
Wenti_form.TreeView1(0).Nodes.Clear
Wenti_form.List1.Clear
inform_data_last_item = 0
inform_type = data_type
Select Case inform_type
Case two_line_value_
For i% = 1 To last_conditions.last_cond(1).two_line_value_no
Call set_inform_list_(i%, _
 set_display_two_line_value(two_line_value(i%), False, 0, False), _
     two_line_value_, i%)
Next i%
inform_caption = LoadResString_(1035, "") & "-" & LoadResString_(1560, "\\1\\" + _
   str(last_conditions.last_cond(1).two_line_value_no))
Case angle2_right
For i% = 1 To last_conditions.last_cond(1).two_angle_value_90_no
n% = two_angle_value_90.av_no(i%).no
Call set_inform_list_(i%, _
 set_display_three_angle_value(angle3_value(n%).data(0), False, 0, False), _
        angle3_value_, n%)
Next i%
inform_caption = LoadResString_(980, "") & "-" & LoadResString_(1560, "\\1\\" + _
   str(last_conditions.last_cond(1).two_angle_value_90_no))
Case two_angle_180_
For i% = 1 To last_conditions.last_cond(1).two_angle_value_180_no
n% = two_angle_value_180.av_no(i%).no
Call set_inform_list_(i%, _
 set_display_three_angle_value(angle3_value(n%).data(0), False, 0, False), _
                 angle3_value_, n%)
Next i%
inform_caption = LoadResString_(985, "") & "-" & LoadResString_(1560, "\\1\\" + _
   str(last_conditions.last_cond(1).two_angle_value_180_no))
Case total_equal_triangle_
For i% = 1 To last_conditions.last_cond(1).total_equal_triangle_no
Call set_inform_list_(i%, _
   set_display_total_equal_triangle(Dtotal_equal_triangle(i%).data(0), False, False), _
      total_equal_triangle_, i%)
Next i%
inform_caption = LoadResString_(1075, "") & "-" & LoadResString_(1560, "\\1\\" + _
   str(last_conditions.last_cond(1).total_equal_triangle_no))
Case point3_on_line_
For i% = 1 To last_conditions.last_cond(1).three_point_on_line_no
Call set_inform_list_(i%, _
   LoadResString_from_inpcond(24, "\\0\\" + m_poi(three_point_on_line(i%).data(0).poi(0)).data(0).data0.name + _
                      "\\1\\" + m_poi(three_point_on_line(i%).data(0).poi(1)).data(0).data0.name + _
                      "\\2\\" + m_poi(three_point_on_line(i%).data(0).poi(2)).data(0).data0.name), _
                      point3_on_line_, i%)
Next i%
inform_caption = LoadResString_(1020, "") & "-" & LoadResString_(1560, "\\1\\" + _
str(last_conditions.last_cond(1).three_point_on_line_no))
Case angle3_value_
For i% = 1 To last_conditions.last_cond(1).three_angle_value_sum_no
 n% = three_angle_value_sum.av_no(i%).no
Call set_inform_list_(i%, _
 set_display_three_angle_value(angle3_value(n%).data(0), False, 0, False), _
        angle3_value_, n%)
Next i%
inform_caption = LoadResString_(975, "") & "-" & LoadResString_(1560, "\\1\\" + _
   str(last_conditions.last_cond(1).three_angle_value_sum_no))
Case sp_polygon4_
n% = 0
For i% = 1 To last_conditions.last_cond(1).tixing_no
If Dpolygon4(Dtixing(i%).data(0).poly4_no).data(0).ty = equal_side_tixing_ Then
n% = n% + 1
Call set_inform_list_(n%, LoadResString_from_inpcond(49, _
  "\\0\\" + m_poi(Dtixing(i%).data(0).poi(0)).data(0).data0.name + _
    "\\1\\" + m_poi(Dtixing(i%).data(0).poi(1)).data(0).data0.name + _
     "\\2\\" + m_poi(Dtixing(i%).data(0).poi(2)).data(0).data0.name + _
      "\\3\\" + m_poi(Dtixing(i%).data(0).poi(3)).data(0).data0.name), _
       polygon_, Dtixing(i%).data(0).poly4_no)
Else
Call set_inform_list_(n%, LoadResString_from_inpcond(48, _
  "\\0\\" + m_poi(Dtixing(i%).data(0).poi(0)).data(0).data0.name + _
   "\\1\\" + m_poi(Dtixing(i%).data(0).poi(1)).data(0).data0.name + _
     "\\2\\" + m_poi(Dtixing(i%).data(0).poi(2)).data(0).data0.name + _
       "\\3\\" + m_poi(Dtixing(i%).data(0).poi(3)).data(0).data0.name), _
       polygon_, Dtixing(i%).data(0).poly4_no)
End If
Next i%
For i% = 1 To last_conditions.last_cond(1).parallelogram_no
n% = n% + 1
Call set_inform_list_(n%, LoadResString_from_inpcond(-11, _
         "\\0\\" + m_poi(Dpolygon4(Dparallelogram(i%).data(0).polygon4_no).data(0).poi(0)).data(0).data0.name + _
         "\\1\\" + m_poi(Dpolygon4(Dparallelogram(i%).data(0).polygon4_no).data(0).poi(1)).data(0).data0.name + _
         "\\2\\" + m_poi(Dpolygon4(Dparallelogram(i%).data(0).polygon4_no).data(0).poi(2)).data(0).data0.name + _
         "\\3\\" + m_poi(Dpolygon4(Dparallelogram(i%).data(0).polygon4_no).data(0).poi(3)).data(0).data0.name), _
           polygon_, Dparallelogram(i%).data(0).polygon4_no)
Next i%
For i% = 1 To last_conditions.last_cond(1).rhombus_no
n% = n% + 1
Call set_inform_list_(n%, LoadResString_from_inpcond(-10, "\\0\\" + _
  m_poi(Dpolygon4(rhombus(i%).data(0).polygon4_no).data(0).poi(0)).data(0).data0.name + _
    "\\1\\" + m_poi(Dpolygon4(rhombus(i%).data(0).polygon4_no).data(0).poi(1)).data(0).data0.name + _
     "\\2\\" + m_poi(Dpolygon4(rhombus(i%).data(0).polygon4_no).data(0).poi(2)).data(0).data0.name + _
      "\\3\\" + m_poi(Dpolygon4(rhombus(i%).data(0).polygon4_no).data(0).poi(3)).data(0).data0.name), _
        polygon_, rhombus(i%).data(0).polygon4_no)
Next i%
For i% = 1 To last_conditions.last_cond(1).long_squre_no
n% = n% + 1
Call set_inform_list_(n%, LoadResString_from_inpcond(-13, _
           "\\0\\" + m_poi(Dpolygon4(Dlong_squre(i%).data(0).polygon4_no).data(0).poi(0)).data(0).data0.name + _
           "\\1\\" + m_poi(Dpolygon4(Dlong_squre(i%).data(0).polygon4_no).data(0).poi(1)).data(0).data0.name + _
           "\\2\\" + m_poi(Dpolygon4(Dlong_squre(i%).data(0).polygon4_no).data(0).poi(2)).data(0).data0.name + _
           "\\3\\" + m_poi(Dpolygon4(Dlong_squre(i%).data(0).polygon4_no).data(0).poi(3)).data(0).data0.name), _
                   polygon_, Dlong_squre(i%).data(0).polygon4_no)

Next i%
For i% = 1 To last_conditions.last_cond(1).squre_no
n% = n% + 1
Call set_inform_list_(n%, LoadResString_(1740, _
        "\\1\\" + m_poi(Dpolygon4(Dsqure(i%).data(0).polygon4_no).data(0).poi(0)).data(0).data0.name + _
        m_poi(Dpolygon4(Dsqure(i%).data(0).polygon4_no).data(0).poi(1)).data(0).data0.name + _
        m_poi(Dpolygon4(Dsqure(i%).data(0).polygon4_no).data(0).poi(2)).data(0).data0.name + _
        m_poi(Dpolygon4(Dsqure(i%).data(0).polygon4_no).data(0).poi(3)).data(0).data0.name), _
               polygon_, Dsqure(i%).data(0).polygon4_no)
Next i%
inform_caption = LoadResString_(1910, "") & "-" & LoadResString_(1560, "\\1\\" + _
              str(n%))
Case angle_value_
For i% = 1 To last_conditions.last_cond(1).angle_value_no
 n% = angle_value.av_no(i%).no
 Call set_inform_list_(i%, _
    set_display_three_angle_value(angle3_value(n%).data(0), False, 0, False), angle3_value_, n%)
Next i%
inform_caption = LoadResString_(995, "") & "-" & LoadResString_(1560, "\\1\\" + _
   str(last_conditions.last_cond(1).angle_value_no))
Case Rangle_ 'If value$ = "30" Then
For i% = 1 To last_conditions.last_cond(1).angle_value_90_no
 n% = angle_value_90.av_no(i%).no
 Call set_inform_list_(i%, _
    set_display_three_angle_value(angle3_value(n%).data(0), False, 0, False), angle3_value_, n%)
Next i%
inform_caption = LoadResString_(1000, "") & "-" & LoadResString_(1560, "\\1\" + _
   str(last_conditions.last_cond(1).angle_value_90_no))
Case relation_
For i% = 1 To last_conditions.last_cond(1).relation_no
Call set_inform_list_(i%, _
 set_display_relation(Drelation(i%), 0, False, 1, 1, False), relation_, i%)
Next i%
inform_caption = LoadResString_(1045, "") & "-" & LoadResString_(1560, "\\1\\" + _
str(last_conditions.last_cond(1).relation_no))
Case dpoint_pair_
For i% = 1 To last_conditions.last_cond(1).dpoint_pair_no
Call set_inform_list_(i%, _
 set_display_point_pair(Ddpoint_pair(i%).data(0).data0, Ddpoint_pair(i%).data(0).record, False, False), _
        dpoint_pair_, i%)
Next i%
inform_caption = LoadResString_(1050, "") & "-" & LoadResString_(1560, "\\1\\" + _
   str(last_conditions.last_cond(1).dpoint_pair_no))
Case line_value_
For i% = 1 To last_conditions.last_cond(1).line_value_no
Call set_inform_list_(i%, _
   set_display_line_value(line_value(i%), False, 0), line_value_, i%)
Next i%
inform_caption = LoadResString_(1030, "") & "-" & LoadResString_(1560, "\\1\\" + _
   str(last_conditions.last_cond(1).line_value_no))
Case point4_on_circle_
For i% = 1 To last_conditions.last_cond(1).four_point_on_circle_no
 Call set_inform_list_(i%, _
  LoadResString_from_inpcond(23, "\\0\\" + m_poi(four_point_on_circle(i%).data(0).poi(0)).data(0).data0.name + _
                     "\\1\\" + m_poi(four_point_on_circle(i%).data(0).poi(1)).data(0).data0.name + _
                     "\\2\\" + m_poi(four_point_on_circle(i%).data(0).poi(2)).data(0).data0.name + _
                     "\\3\\" + m_poi(four_point_on_circle(i%).data(0).poi(3)).data(0).data0.name), _
                     point4_on_circle_, i%)
Next i%
inform_caption = LoadResString_(1060, "") & "-" & LoadResString_(1560, "\\1\\" + _
str(last_conditions.last_cond(1).four_point_on_circle_no))
Case eline_
For i% = 1 To last_conditions.last_cond(1).eline_no
Call set_inform_list_(i%, _
set_display_eline(Deline(i%).data(0), False, False), eline_, i%)
Next i%
For i% = 1 To last_conditions.last_cond(1).mid_point_no
Call set_inform_list_(last_conditions.last_cond(1).eline_no + i%, _
set_display_mid_point(Dmid_point(i%), 1, False, False), midpoint_, i%)
Next i%
inform.Caption = LoadResString_(1040, "") & "-" & LoadResString_(1560, "\\1\\" + _
str(last_conditions.last_cond(1).eline_no + last_conditions.last_cond(1).mid_point_no))
Case eangle_
For i% = 1 To last_conditions.last_cond(1).eangle_no
 n% = Deangle.av_no(i%).no
Call set_inform_list_(i%, _
   set_display_three_angle_value(angle3_value(n%).data(0), False, 0, False), three_angle_value_, n%)
Next i%
inform_caption = LoadResString_(990, "") & "-" & LoadResString_(1560, "\\1\\" + _
   str(last_conditions.last_cond(1).eangle_no))
Case epolygon_
For i% = 1 To last_conditions.last_cond(1).epolygon_no
Call set_inform_list_(i%, set_display_Epolygon(epolygon(i%).data(0), False, 0, 0), epolygon_, i%)
Next i%
inform_caption = LoadResString_(1920, "") & "-" & LoadResString_(1560, "\\1\\" + _
                   str(last_conditions.last_cond(1).epolygon_no))
Case similar_triangle_
For i% = 1 To last_conditions.last_cond(1).similar_triangle_no
Call set_inform_list_(i%, "(" + Trim(str(i%)) + ")" + _
 set_display_similar_triangle(Dsimilar_triangle(i%).data(0), False, False), similar_triangle_, i%)
Next i%
inform_caption = LoadResString_(1080, "") & "-" & LoadResString_(1560, "\\1\\" + _
   str(last_conditions.last_cond(1).similar_triangle_no))
Case angle_relation_
k% = 1
 For i% = 1 To last_conditions.last_cond(1).angle_relation_no
  n% = angle_relation.av_no(i%).no
   Call set_inform_list_(k%, set_display_three_angle_value(angle3_value(n%).data(0), False, 0, False), _
                              angle_relation_, n%)
   k% = k% + 1
 Next i%
 inform_caption = LoadResString_(970, "") & "-" & LoadResString_(1560, "\\1\\" + _
   str(last_conditions.last_cond(1).angle_relation_no))
Case area_of_circle_
  For i% = 1 To last_conditions.last_cond(1).area_of_circle_no
     Call set_inform_list_(i%, "(" + Trim(str(i%)) + ")" + LoadResString_(1795, "\\1\" + LoadResString_(1410, "") + _
       m_poi(m_Circ(area_of_circle(i%).data(0).circ).data(0).data0.center).data(0).data0.name + _
       "\\2\\" + display_string_(area_of_circle(i%).data(0).value, 1)), area_of_circle_, i%)
  Next i%
  inform_caption = LoadResString_(1065, "") & "-" & LoadResString_(1560, "\\1\\" + _
   str(last_conditions.last_cond(1).area_of_circle_no))
Case area_of_polygon_
n% = 0
For i% = 1 To last_conditions.last_cond(1).area_of_element_no
If area_of_element(i%).data(0).element.ty = polygon_ Then
n% = n% + 1
Call set_inform_list_(n%, _
      set_area_element_display_string(area_of_element(i%).data(0), 0, False), _
        area_of_element_, i%)
End If
Next i%
inform_caption = LoadResString_(595, "\\1\\" + LoadResString_(1910, "")) & "-" & LoadResString_(1560, "\\1\\" + _
                      str(n%))
Case area_of_triangle_
inform_type = area_of_element_
n% = 0
For i% = 1 To last_conditions.last_cond(1).area_of_element_no
If area_of_element(i%).data(0).element.ty = triangle_ Then
n% = n% + 1
Call set_inform_list_(n%, "(" + Trim(str(n%)) + ")" + _
      set_area_element_display_string(area_of_element(i%).data(0), True, False), area_of_element_, i%)
End If
Next i%
inform_caption = LoadResString_(1085, "") & "-" & LoadResString_(1560, "\\1\\" + _
               str(n%))
Case paral_
 For i% = 1 To last_conditions.last_cond(1).paral_no
  Call set_inform_list_(i%, set_display_paral(Dparal(i%).data(0).data0, False, 0, False), paral_, i%)
 Next i%
   inform_caption = LoadResString_(1010, "") & "-" & LoadResString_(1560, "\\1\\" + _
    str(last_conditions.last_cond(1).paral_no))
Case two_angle_value_sum_ 'angle2_right
 For i% = 1 To last_conditions.last_cond(1).two_angle_value_sum_no
  n% = two_angle_value_sum.av_no(i%).no
   Call set_inform_list_(i%, set_display_three_angle_value(angle3_value(n%).data(0), False, 0, False), _
                              three_angle_value_, n%)
  Next i%
   inform_caption = LoadResString_(965, "") & "-" & LoadResString_(1560, "\\1\\" + _
    str(last_conditions.last_cond(1).two_angle_value_sum_no))
Case verti_
   For i% = 1 To last_conditions.last_cond(1).verti_no
     Call set_inform_list_(i%, set_display_verti(Dverti(i%).data(0), False, False), _
                              verti_, i%)
   Next i%
    inform_caption = LoadResString_(1015, "") & "-" & LoadResString_(1560, "\\1\\" + _
     str(last_conditions.last_cond(1).verti_no))
End Select
Wenti_form.Picture2.Cls
Wenti_form.Picture2.CurrentX = 0
Wenti_form.Picture2.CurrentY = 0
Wenti_form.Picture2.Print LoadResString_(955, "") & ":" & inform_caption
   Wenti_form.SSTab1.Tab = 1
   Wenti_form.SSTab1.Caption = LoadResString_(955, "")
   SSTab1_name_type = 1
     MDIForm1.StatusBar1.Panels(1).text = LoadResString_(505, "")
End Sub
Private Sub set_inform_list_(i%, t_s As String, ty As Byte, no%)
If inform_data_last_item Mod 10 = 0 Then
ReDim Preserve inform_data_base(inform_data_last_item + 10) As condition_type
End If
inform_data_base(inform_data_last_item).ty = ty
inform_data_base(inform_data_last_item).no = no%
Wenti_form.List1.AddItem "(" + Trim(str(i%)) + ")" + t_s
inform_data_last_item = inform_data_last_item + 1
End Sub

Public Sub delete_inform()
      MDIForm1.StatusBar1.Panels(1).text = "" '_
      Call draw_inform(0, 0, 0)
      MDIForm1.Timer1.Enabled = False
End Sub

Public Sub draw_inform(w_no%, ty As Integer, no%)
 If inform_condition_data.data.ty <> ty Or _
       inform_condition_data.data.no <> no% Then
       MDIForm1.time11_display_type = 0
         Call draw_picture_for_inform_(0)
              inform_condition_data.data.ty = ty
              inform_condition_data.data.no = no%
              inform_condition_data.wenti_no = w_no%
              MDIForm1.time11_display_type = inform_
              MDIForm1.Timer1.Enabled = True
  Call draw_picture_for_inform_(1)
  End If
End Sub

Public Sub draw_picture_for_inform_(display_or_delete As Byte) 'ty=0 ÊäÈëÓï¾ä
Dim i%, j%, p%, no%
Dim ty As Byte
Dim draw_color As Long
Dim tp(3) As Integer
Dim A_data As angle_data_type
If display_or_delete = 1 Then
draw_color = QBColor(13) 'co0
Else
draw_color = QBColor(9)
End If
ty = inform_condition_data.data.ty
no% = inform_condition_data.data.no
If inform_condition_data.wenti_no > 0 Then
 If display_or_delete = 0 Then
   Call C_display_wenti.m_display_string.item(inform_condition_data.wenti_no). _
         display_m_input_condi_(1, 0, 0, 1)
 Else
   Call C_display_wenti.m_display_string.item(inform_condition_data.wenti_no). _
         display_m_input_condi_(0, 0, 0, 1)
 End If
End If
If ty = 0 Then
   Exit Sub
ElseIf ty = midpoint_ Then
        Call draw_line_for_inform(Dmid_point(no%).data(0).data0.poi(0), _
              Dmid_point(no%).data(0).data0.poi(2), display_or_delete)
ElseIf ty = equal_arc_ Then '-24
    Call draw_arc_for_inform(arc(equal_arc(no%).data(0).arc(0)).data(0).poi(0), _
                            arc(equal_arc(no%).data(0).arc(0)).data(0).poi(0), draw_color)
    Call draw_arc_for_inform(arc(equal_arc(no%).data(0).arc(1)).data(0).poi(0), _
                            arc(equal_arc(no%).data(0).arc(1)).data(0).poi(0), draw_color)
ElseIf ty = arc_value_ Then
   Call draw_arc_for_inform(arc(arc_value(no%).data(0).arc).data(0).poi(0), _
                            arc(arc_value(no%).data(0).arc).data(0).poi(0), draw_color)
ElseIf ty = paral_ Then '2
        Call draw_line_for_inform(m_lin(Dparal(no%).data(0).data0.line_no(0)).data(0).data0.poi(0), _
              m_lin(Dparal(no%).data(0).data0.line_no(0)).data(0).data0.poi(1), display_or_delete)
        Call draw_line_for_inform(m_lin(Dparal(no%).data(0).data0.line_no(1)).data(0).data0.poi(0), _
              m_lin(Dparal(no%).data(0).data0.line_no(1)).data(0).data0.poi(1), display_or_delete)
ElseIf ty = verti_ Then '3
        Call draw_line_for_inform(m_lin(Dverti(no%).data(0).line_no(0)).data(0).data0.poi(0), _
              m_lin(Dverti(no%).data(0).line_no(0)).data(0).data0.poi(1), display_or_delete)
        Call draw_line_for_inform(m_lin(Dverti(no%).data(0).line_no(1)).data(0).data0.poi(0), _
              m_lin(Dverti(no%).data(0).line_no(1)).data(0).data0.poi(1), display_or_delete)
ElseIf ty = circle_ Then
       Call draw_circle_for_inform(no%, display_or_delete)
ElseIf ty = line_ Then
              Call draw_line_for_inform(m_lin(no%).data(0).data0.poi(0), _
                               m_lin(no%).data(0).data0.poi(1), display_or_delete)
ElseIf ty = polygon_ Then
              Call draw_epolygon_for_inform(epolygon(no%).data(0), draw_color)
ElseIf ty = tangent_line_ Then
        Call draw_circle_for_inform(tangent_line(no%).data(0).circ(0), display_or_delete)
        If tangent_line(no%).data(0).circ(1) > 0 Then
        Call draw_circle_for_inform(tangent_line(no%).data(0).circ(0), display_or_delete)
        End If
        Call draw_line_for_inform(m_lin(tangent_line(no%).data(0).line_no).data(0).data0.poi(0), _
              m_lin(Dverti(no%).data(0).line_no(1)).data(0).data0.poi(1), display_or_delete)
ElseIf ty = eline_ Then
        Call draw_line_for_inform(Deline(no%).data(0).data0.poi(0), _
              Deline(no%).data(0).data0.poi(1), display_or_delete)
        Call draw_line_for_inform(Deline(no%).data(0).data0.poi(2), _
              Deline(no%).data(0).data0.poi(3), display_or_delete)
ElseIf ty = three_angle_value_ Then
        Call draw_angle_for_inform_(angle3_value(no%).data(0).data0.angle(0), display_or_delete)
        Call draw_angle_for_inform_(angle3_value(no%).data(0).data0.angle(1), display_or_delete)
        Call draw_angle_for_inform_(angle3_value(no%).data(0).data0.angle(2), display_or_delete)
ElseIf ty = tangent_line_ Then
        Call draw_line_for_inform(m_lin(tangent_line(no%).data(0).line_no).data(0).data0.poi(0), _
                                       m_lin(tangent_line(no%).data(0).line_no).data(0).data0.poi(1), display_or_delete)
        Call draw_circle_for_inform(tangent_line(no%).data(0).circ(0), display_or_delete)
           If tangent_line(no%).data(0).circ(1) > 0 Then
           Call draw_circle_for_inform(tangent_line(no%).data(0).circ(1), display_or_delete)
           End If
ElseIf ty = point4_on_circle_ Then
        Call draw_circle_for_inform(four_point_on_circle(no%).data(0).circ, _
                                     display_or_delete)
ElseIf ty = point3_on_line_ Then
        Call draw_line_for_inform(three_point_on_line(no%).data(0).poi(0), _
                                    three_point_on_line(no%).data(0).poi(2), display_or_delete)
ElseIf ty = angle3_value_ Then
        Call draw_angle_for_inform_(angle3_value(no%).data(0).data0.angle(0), display_or_delete)
        Call draw_angle_for_inform_(angle3_value(no%).data(0).data0.angle(1), display_or_delete)
        Call draw_angle_for_inform_(angle3_value(no%).data(0).data0.angle(2), display_or_delete)
Else
End If
End Sub

Public Sub draw_angle_for_inform_(angle_no%, display_or_delete As Byte)
Dim tp(2) As Integer
If angle_no% > 0 Then
tp(0) = m_lin(angle(angle_no%).data(0).line_no(0)).data(0).data0.poi(angle(angle_no%).data(0).te(0))
tp(1) = angle(angle_no%).data(0).poi(1)
tp(2) = m_lin(angle(angle_no%).data(0).line_no(1)).data(0).data0.poi(angle(angle_no%).data(0).te(1))
        Call draw_line_for_inform(tp(0), tp(1), display_or_delete)
        Call draw_line_for_inform(tp(1), tp(2), display_or_delete)
End If
'Call draw_angle_for_inform(tp(0), tp(1), tp(2), co, fillstyle)
End Sub

Public Sub draw_angle_for_inform(p1%, p2%, p3%, co As Long, fillstyle As Byte)
Dim i%
Dim r As Long
Dim p(2) As POINTAPI
If line_width < 2 Then
Draw_form.DrawWidth = 2
End If
Draw_form.FillColor = co
Draw_form.fillstyle = 1
 Call Drawline(Draw_form, co, 0, _
      m_poi(p2%).data(0).data0.coordinate, _
        m_poi(p1%).data(0).data0.coordinate, 0)
 Call Drawline(Draw_form, co, 0, _
      m_poi(p2%).data(0).data0.coordinate, _
        m_poi(p3%).data(0).data0.coordinate, 0)
      p(1).X = p(1).X + m_poi(p2%).data(0).data0.coordinate.X
      p(1).Y = p(1).Y + m_poi(p2%).data(0).data0.coordinate.Y
      r = sqr((m_poi(p1%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) ^ 2 + _
           (m_poi(p1%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y) ^ 2)
      p(0).X = m_poi(p2%).data(0).data0.coordinate.X + _
        20 * (m_poi(p1%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) / r
      p(0).Y = m_poi(p2%).data(0).data0.coordinate.Y + _
        20 * (m_poi(p1%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y) / r
      p(1).X = p(1).X + p(0).X
      p(1).Y = p(1).Y + p(0).Y
 Call Drawline(Draw_form, co, 0, _
      m_poi(p2%).data(0).data0.coordinate, p(0), 0)
'********
      r = sqr((m_poi(p3%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) ^ 2 + _
           (m_poi(p3%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y) ^ 2)
      p(2).X = m_poi(p2%).data(0).data0.coordinate.X + _
        20 * (m_poi(p3%).data(0).data0.coordinate.X - m_poi(p2%).data(0).data0.coordinate.X) / r
      p(2).Y = m_poi(p2%).data(0).data0.coordinate.Y + _
        20 * (m_poi(p3%).data(0).data0.coordinate.Y - m_poi(p2%).data(0).data0.coordinate.Y) / r
      p(1).X = p(1).X + p(2).X
      p(1).Y = p(1).Y + p(2).Y
 Call Drawline(Draw_form, co, 0, _
      m_poi(p2%).data(0).data0.coordinate, p(2), 0)
 Draw_form.DrawWidth = 1
 Call Drawline(Draw_form, co, 0, p(0), p(2), 0)
p(1).Y = p(1).Y / 3
p(1).X = p(1).X / 3
Draw_form.fillstyle = fillstyle
Call FloodFill(Draw_form.hdc, p(1).X, _
       p(1).Y, co)
Draw_form.fillstyle = line_width
End Sub
Public Sub draw_circle_for_inform(circle_no%, display_or_delete)
Draw_form.fillstyle = 1
If line_width < 2 Then
Draw_form.DrawWidth = 2
End If
If circle_no% = 0 Then
 Exit Sub
Else
If m_Circ(circle_no%).data(0).circle_type = conclusion Then
   If display_or_delete = 0 Then
    Call C_display_picture.redraw_circle(circle_no%, True)
   Else
    Call C_display_picture.draw_circle(circle_no%, 0, 0, 0)
   End If
Else
 If display_or_delete = 0 Then
  Call C_display_picture.redraw_circle(circle_no%, True)
 Else
  Call C_display_picture.draw_circle(circle_no%, 0, 0, 12)
 End If
End If
End If
'm_picture.
'If c.center = 0 Then
'Call draw_three_point_circle(m_poi(c.in_point(1)).data(0).data0.coordinate.X, _
 m_poi(c.in_point(1)).data(0).data0.coordinate.Y, m_poi(c.in_point(2)).data(0).data0.coordinate.X, _
  m_poi(c.in_point(2)).data(0).data0.coordinate.Y, m_poi(c.in_point(3)).data(0).data0.coordinate.X, _
   m_poi(c.in_point(3)).data(0).data0.coordinate.Y, co, 0, 0, 0, display, 1)
'Else
'Draw_form.Circle (m_poi(c.center).data(0).data0.coordinate.X, m_poi _
          (c.center).data(0).data0.coordinate.Y), c.radii, co
'End If
'Draw_form.DrawWidth = line_width
End Sub
Public Sub draw_arc_for_inform(p1%, p2%, co As Long)
Dim i%, c%
Draw_form.fillstyle = 1
If line_width < 2 Then
Draw_form.DrawWidth = 2
End If
For i% = 1 To last_conditions.last_cond(1).arc_no
 If is_same_two_point(p1%, p2%, arc(i%).data(0).poi(0), arc(i%).data(0).poi(1)) Then
  c% = arc(i%).data(0).cir
   GoTo draw_arc_for_inform_mark0
 End If
Next i%
draw_arc_for_inform_mark0:
 Draw_form.Circle (m_poi(m_Circ(c%).data(0).data0.center).data(0).data0.coordinate.X, m_poi _
          (m_Circ(c%).data(0).data0.center).data(0).data0.coordinate.Y), m_Circ(c%).data(0).data0.radii, co
Draw_form.DrawWidth = line_width
End Sub
Public Sub fill_color_circle_for_inform(c As circle_data0_type, co As Long)
Draw_form.fillstyle = 4
If line_width < 2 Then
Draw_form.DrawWidth = 2
End If
If c.center = 0 Then
Call draw_three_point_circle(m_poi(c.in_point(1)).data(0).data0.coordinate.X, _
 m_poi(c.in_point(1)).data(0).data0.coordinate.Y, m_poi(c.in_point(2)).data(0).data0.coordinate.X, _
  m_poi(c.in_point(2)).data(0).data0.coordinate.Y, m_poi(c.in_point(3)).data(0).data0.coordinate.X, _
   m_poi(c.in_point(3)).data(0).data0.coordinate.Y, co, 0, 0, 0, display, 1)
Else
Draw_form.Circle (m_poi(c.center).data(0).data0.coordinate.X, m_poi _
          (c.center).data(0).data0.coordinate.Y), c.radii, co
End If
Draw_form.DrawWidth = line_width
End Sub


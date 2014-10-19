Attribute VB_Name = "change"
Option Explicit
Type subs_angle3_value_type
no As Integer
A3_value As angle3_value_data0_type
End Type
Global subs_angle3_value() As subs_angle3_value_type
Global last_subs_angle3_value As Integer

Public Sub simple_record_of_dbase(ty1 As Byte, n1%, ty2 As Byte, n2%, ty As Byte)
Dim i%, j%
Dim replace_n%, n%
Dim cond_ty As Byte
Dim re_ty As Byte
Dim re(1) As total_record_type
If ty = 0 Then
   re_ty = ty2
   cond_ty = ty1
   replace_n% = n2%
   n% = n1%
Else
Call record_no(ty1, n1%, re(0), False, 0, 0)
Call record_no(ty2, n2%, re(1), False, 0, 0)
If re(0).record_data.data0.condition_data.level < re(1).record_data.data0.condition_data.level Then
   re_ty = ty1
   cond_ty = ty2
   replace_n% = n1%
   n% = n2%
Else
   re_ty = ty2
   cond_ty = ty1
   replace_n% = n2%
   n% = n1%
End If
End If
For i% = 1 To last_conditions.last_cond(1).verti_no '24
 Call simple_record_by_replace(Dverti(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).epolygon_no
 Call simple_record_by_replace(epolygon(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).tixing_no
 Call simple_record_by_replace(Dtixing(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).rhombus_no
 Call simple_record_by_replace(rhombus(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)

Next i%
For i% = 1 To last_conditions.last_cond(1).long_squre_no
 Call simple_record_by_replace(Dlong_squre(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)

Next i%
'For i% = 1 To last_angle_value
' Call simple_record_by_replace(angle_value(i%).record, cond_ty, replace_n%, n%)
'Next i%
For i% = 1 To last_conditions.last_cond(1).dpoint_pair_no  '3
 Call simple_record_by_replace(Ddpoint_pair(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).area_relation_no  '4
 Call simple_record_by_replace(Darea_relation(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%

'For i = 1 To last_Eangle '6
' Call simple_record_by_replace(Deangle(i%).record, cond_ty, replace_n%, n%)
'Next i%
For i% = 1 To last_conditions.last_cond(1).mid_point_line_no
 Call simple_record_by_replace(mid_point_line(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).eline_no '8
 Call simple_record_by_replace(Deline(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).four_point_on_circle_no '9
 Call simple_record_by_replace(four_point_on_circle(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).three_point_on_circle_no '9
 Call simple_record_by_replace(three_point_on_circle(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%

'For i% = 1 To last_angle_relation '10
' Call simple_record_by_replace(angle_relation(i%).record, cond_ty, replace_n%, n%)
'Next i%
For i% = 1 To last_conditions.last_cond(1).mid_point_no '12
 Call simple_record_by_replace(Dmid_point(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).paral_no
 Call simple_record_by_replace(Dparal(i%).data(0).data0.record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).parallelogram_no '14
 Call simple_record_by_replace(Dparallelogram(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).relation_no '16
 Call simple_record_by_replace(Drelation(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).triangle_no
  If re_ty = midpoint_ And cond_ty = midpoint_ Then
   For j% = 0 To 2
    If triangle(i%).data(0).midpoint_no(j%) = n% Then
     triangle(i%).data(0).midpoint_no(j%) = replace_n%
    End If
   Next j%
  ElseIf re_ty = line_value_ And cond_ty = line_value_ Then
   'For j% = 0 To 2
    'If triangle(i%).data(0).angle_value(j%) = n% Then
     'triangle(i%).data(0).angle_value(j%) = replace_n%
    'End If
   'Next j%
  ElseIf re_ty = relation_ And cond_ty = relation_ Then
   For j% = 0 To 2
    If triangle(i%).data(0).relation_no(j%, 0).no = n% And _
         triangle(i%).data(0).relation_no(j%, 0).ty = relation_ Then
     triangle(i%).data(0).relation_no(j%, 0).no = replace_n%
    End If
 Next j%
  ElseIf re_ty = angle3_value_ And cond_ty = angle3_value_ Then
   'For j% = 0 To 2
    'If triangle(i%).data(0).angle_value(j%) = n% Then
     'triangle(i%).data(0).angle_value(j%) = replace_n%
    'End If
   'Next j%
 ElseIf re_ty = area_of_element_ And cond_ty = area_of_element_ Then
  If triangle(i%).data(0).area_no = n% Then
   triangle(i%).data(0).area_no = replace_n%
  End If
 End If
Next i%
For i% = 1 To last_conditions.last_cond(1).similar_triangle_no '17
 Call simple_record_by_replace(Dsimilar_triangle(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).total_equal_triangle_no  '18
 Call simple_record_by_replace(Dtotal_equal_triangle(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).three_point_on_line_no '20
 Call simple_record_by_replace(three_point_on_line(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
'For i% = 1 To last_conditions.last_cond(1).two_point_conset_no '20
 'Call simple_record_by_replace(three_point_on_line(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
'Next i%
For i% = 1 To last_conditions.last_cond(1).two_line_value_no '21
 Call simple_record_by_replace(two_line_value(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
'For i% = 1 To Last_two_angle_value '22
' Call simple_record_by_replace(Two_angle_value(i%).record, cond_ty, replace_n%, n%)
'Next i%
For i% = 1 To last_conditions.last_cond(1).arc_value_no '25
 Call simple_record_by_replace(arc_value(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).equal_arc_no '26
 Call simple_record_by_replace(equal_arc(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).line_value_no '33
 Call simple_record_by_replace(line_value(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).tangent_line_no  '34
 Call simple_record_by_replace(tangent_line(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
For i% = 1 To last_conditions.last_cond(1).area_of_element_no '35
 Call simple_record_by_replace(area_of_element(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
Next i%
'For i% = 1 To last_conditions.last_cond(1).equal_area_triangle_no '36
' Call simple_record_by_replace(equal_area_triangle(i%).data(0).record, re_ty, cond_ty, replace_n%, n%)
'Next i%

End Sub
Public Function simple_dbase_for_angle(ByVal no%, ByVal replace_no%, ByVal ty_ As Byte, re As total_record_type) As Byte
'no% 原角号,replace_no%替换角号,deg%零角或平角
Dim i%, j%, tn%, k%, l%, m%
Dim n_(8) As Integer
Dim n1_(8) As Integer
Dim tn_() As Integer
Dim last_tn%
Dim ty As Byte
Dim t_a_v(2) As angle3_value_data0_type
Dim temp_record As total_record_type
Dim a3_v_type As Byte
'**********
ty = 0
For i% = 0 To 3
 If conclusion_data(i%).ty = angle3_value_ Then
    If con_angle3_value(i%).data(0).data0.angle(0) = no% Then
         con_angle3_value(i%).data(0).data0.angle(0) = replace_no%
            If ty_ = 1 Then
              con_angle3_value(i%).data(0).data0.para(0) = time_string("-1", _
                 con_angle3_value(i%).data(0).data0.para(0), True, False)
              con_angle3_value(i%).data(0).data0.value = add_string( _
                 con_angle3_value(i%).data(0).data0.value, time_string( _
                  con_angle3_value(i%).data(0).data0.para(0), "180", False, False), True, False)
            End If
       ty = 1
    ElseIf con_angle3_value(i%).data(0).data0.angle(1) = no% Then
        con_angle3_value(i%).data(0).data0.angle(1) = replace_no%
            If ty_ = 1 Then
              con_angle3_value(i%).data(0).data0.para(1) = time_string("-1", _
                 con_angle3_value(i%).data(0).data0.para(1), True, False)
              con_angle3_value(i%).data(0).data0.value = add_string( _
                 con_angle3_value(i%).data(0).data0.value, time_string( _
                  con_angle3_value(i%).data(0).data0.para(1), "180", False, False), True, False)
            End If
       ty = 1
     ElseIf con_angle3_value(i%).data(0).data0.angle(2) = no% Then
        con_angle3_value(i%).data(0).data0.angle(2) = replace_no%
            If ty_ = 1 Then
              con_angle3_value(i%).data(0).data0.para(2) = time_string("-1", _
                 con_angle3_value(i%).data(0).data0.para(2), True, False)
              con_angle3_value(i%).data(0).data0.value = add_string( _
                 con_angle3_value(i%).data(0).data0.value, time_string( _
                  con_angle3_value(i%).data(0).data0.para(2), "180", False, False), True, False)
            End If
       ty = 1
      End If
    If ty = 1 Then
      con_angle3_value(i%).data(0).data0 = simple_three_angle_value0(con_angle3_value(i%).data(0).data0, 1)
      con_angle3_value(i%).data(0).record = re.record_data
    End If
 End If
Next i%
'On Error GoTo simple_abase_for_angle_error
For j% = 0 To 2
 t_a_v(2).angle(j%) = no%
  t_a_v(2).angle((j% + 1) Mod 3) = -1
 Call search_for_three_angle_value(t_a_v(2), j%, n_(0), 1)
  t_a_v(2).angle((j% + 1) Mod 3) = 30000
 Call search_for_three_angle_value(t_a_v(2), j%, n_(1), 1)
  last_tn% = 0
 For i% = n_(0) + 1 To n_(1)
  last_tn% = last_tn% + 1
   ReDim Preserve tn_(last_tn%) As Integer
    tn_(last_tn%) = angle3_value(i%).data(0).record.data1.index.i(j%)
 Next i%
 For i% = 1 To last_tn%
 k% = tn_(i%)
  For l% = 1 To last_subs_angle3_value
    If subs_angle3_value(l%).no = k% Then
     GoTo simple_dbase_for_angle_out1
    End If
  Next l%
  l% = 0
simple_dbase_for_angle_out1:
If l% = 0 Then
 If last_subs_angle3_value Mod 10 = 0 Then
  ReDim Preserve subs_angle3_value(last_subs_angle3_value + 10) As subs_angle3_value_type
 End If
  last_subs_angle3_value = last_subs_angle3_value + 1
  l% = last_subs_angle3_value
  subs_angle3_value(l%).no = k%
End If
   subs_angle3_value(l%).A3_value.angle(j%) = replace_no%
      If ty_ = 1 Then
       subs_angle3_value(l%).A3_value.para(j%) = time_string(subs_angle3_value(l%).A3_value.para(j%), "-1", True, False)
        subs_angle3_value(l%).A3_value.value = add_string(subs_angle3_value(l%).A3_value.value, _
              time_string("180", subs_angle3_value(l%).A3_value.para(j%), False, False), True, False)
      End If
Next i%
Next j%
simple_abase_for_angle_error:
End Function
Public Function remove_record(ByVal record_type As Byte, ByVal no%, ty As Byte) As Boolean
Dim i%, j%, k% 'ty=0 不参加推导
Dim n_(8) As Integer
Dim is_true As Boolean
'On Error GoTo remove_record_error
remove_record = True
Select Case record_type
Case item0_
 For i% = 0 To 2
 If search_for_item0(item0(no%).data(0), i%, n_(i%), 1) = False Then
  remove_record = False
 End If
 Next i%
 If remove_record Then
  For i% = 0 To 2
   For j% = n_(i%) + 1 To 2 Step -1
   item0(j%).data(0).index(i%) = _
    item0(j% - 1).data(0).index(i%)
   Next j%
   item0(1).data(0).index(i%) = 0
 Next i%
 last_conditions.last_cond(0).item0_no = last_conditions.last_cond(0).item0_no + 1
 End If
 '*****************
Case total_angle_
remove_record = True
 For i% = 0 To 1
 If search_for_total_angle(T_angle(no%).data(0), n_(0), i%, 1) = False Then
     remove_record = False
 End If
 Next i%
 If remove_record Then
  For i% = 0 To 1
   For j% = n_(0) + 1 To 2 Step -1
    T_angle(j%).data(0).index(i%) = _
     T_angle(j% - 1).data(0).index(i%)
   Next j%
    T_angle(1).data(0).index(i%) = 0
Next i%
 last_conditions.last_cond(0).total_angle_no = last_conditions.last_cond(0).total_angle_no + 1
End If
'********************
Case angle_
If ty = 0 Then
angle(no%).data(0).no_reduce = 255
End If
remove_record = True
  For i% = 0 To 1
  If search_for_angle(angle(no%).data(0), n_(i%), i%, 1) = False Then
   remove_record = False
  End If
  Next i%
  If remove_record Then
   For i% = 0 To 1
   For j% = n_(i%) + 1 To 2 Step -1
    angle(j%).data(0).index(i%) = _
     angle(j% - 1).data(0).index(i%)
   Next j%
    angle(1).data(0).index(i%) = 0
 Next i%
   last_conditions.last_cond(0).angle_no = last_conditions.last_cond(0).angle_no + 1
  End If
'****************
Case line_from_two_point_
If ty = 0 Then
Dtwo_point_line(no%).data(0).no_reduce = 255
End If
If search_for_two_point_line(Dtwo_point_line(no%).poi(0), _
       Dtwo_point_line(no%).poi(1), n_(0), 1) Then
For j% = n_(0) + 1 To 2 Step -1
Dtwo_point_line(j%).data(0).index = _
 Dtwo_point_line(j% - 1).data(0).index
Next j%
Dtwo_point_line(1).data(0).index = 0
last_conditions.last_cond(0).line_from_two_point_no = _
    last_conditions.last_cond(0).line_from_two_point_no + 1
End If
'***************************
'*************************
Case dpoint_pair_
If ty = 0 Then
Ddpoint_pair(no%).record_.no_reduce = 255
End If
For i% = 0 To 6
If search_for_point_pair(Ddpoint_pair(no%).data(0).data0, i%, n_(i%), 1) = False Then
 remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 6
For j% = n_(i%) + 1 To 2 Step -1
Ddpoint_pair(j%).data(0).record.data1.index.i(i%) = _
 Ddpoint_pair(j% - 1).data(0).record.data1.index.i(i%)
Next j%
Ddpoint_pair(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).dpoint_pair_no = _
     last_conditions.last_cond(0).dpoint_pair_no + 1
Ddpoint_pair(no%).data(0).record.data1.is_removed = True
End If
'****************************************************
Case area_relation_
If ty = 0 Then
Darea_relation(no%).record_.no_reduce = 255
End If
For i% = 0 To 1
If search_for_area_relation(Darea_relation(no%).data(0), i%, n_(i%), 1) = False Then
 remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 1
For j% = n_(i%) + 1 To 1 Step -1
Darea_relation(j%).data(0).record.data1.index.i(i%) = _
Darea_relation(j% - 1).data(0).record.data1.index.i(i%)
Next j%
Darea_relation(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).area_relation_no = _
     last_conditions.last_cond(0).area_relation_no + 1
''last_Dpoint = last_dpoint_for_aid '5
'last_mid_point_line = last_mid_point_line_for_aid '7
'last_eline = last_eline_for_aid '8
Darea_relation(no%).data(0).record.data1.is_removed = True
End If
'*********************************************************
Case eline_
If ty = 0 Then
Deline(no%).record_.no_reduce = 255
End If
For i% = 0 To 3
If search_for_eline(Deline(no%).data(0).data0, i%, n_(i%), 1) = False Then
 remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 3
For j% = n_(i%) + 1 To 2 Step -1
Deline(j%).data(0).record.data1.index.i(i%) = _
 Deline(j% - 1).data(0).record.data1.index.i(i%)
Next j%
Deline(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).eline_no = _
     last_conditions.last_cond(0).eline_no + 1

'last_Four_point_on_circle = last_four_point_on_circle_for_aid '9
Deline(no%).data(0).record.data1.is_removed = True
End If
'************************************************
Case point4_on_circle_
If ty = 0 Then
four_point_on_circle(no%).record_.no_reduce = 255
End If
For i% = 0 To 3
If search_for_four_point_on_circle(four_point_on_circle(no%).data(0), 1, i%, n_(i%), 1) = False Then
remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 3
For j% = n_(i%) + 1 To 2 Step -1
four_point_on_circle(j%).data(0).record.data1.index.i(i%) = _
 four_point_on_circle(j% - 1).data(0).record.data1.index.i(i%)
Next j%
four_point_on_circle(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).four_point_on_circle_no = _
     last_conditions.last_cond(0).four_point_on_circle_no + 1
four_point_on_circle(no%).data(0).record.data1.is_removed = True
End If
'****************************************************
Case midpoint_
If ty = 0 Then
Dmid_point(no%).record_.no_reduce = 255
End If
For i% = 0 To 2
If search_for_mid_point(Dmid_point(no%).data(0).data0, i%, n_(i%), 1) = False Then '5.7
 remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 2
For j% = n_(i%) + 1 To 1 Step -1
Dmid_point(j%).data(0).record.data1.index.i(i%) = _
 Dmid_point(j% - 1).data(0).record.data1.index.i(i%)
Next j%
Dmid_point(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).new_point_no = _
    last_conditions.last_cond(0).new_point_no + 1
'last_paral = last_paral_for_aid '13
Dmid_point(no%).data(0).record.data1.is_removed = True
End If
'*******************************************************
Case paral_
If ty = 0 Then
Dparal(no%).record_.no_reduce = 255
End If
For i% = 0 To 1
If search_for_paral(Dparal(no%).data(0).data0, i%, n_(i%), 1) = False Then
remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 1
For j% = n_(i%) + 1 To 2 Step -1
Dparal(j%).data(0).data0.record.data1.index.i(i%) = _
 Dparal(j% - 1).data(0).data0.record.data1.index.i(i%)
Next j%
Dparal(1).data(0).data0.record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).paral_no = _
   last_conditions.last_cond(0).paral_no + 1
'last_parallelogram = last_parallelogram_for_aid '14
Dparal(no%).data(0).data0.record.data1.is_removed = True
End If
'**********************************************************
Case parallelogram_
If ty = 0 Then
Dparallelogram(no%).record_.no_reduce = 255
End If
If search_for_parallelogram(Dparallelogram(no%).data(0).polygon4_no, n_(0), 1) Then
For j% = n_(i%) + 1 To 2 Step -1
Dparallelogram(j%).data(0).record.data1.index.i(0) = _
 Dparallelogram(j% - 1).data(0).record.data1.index.i(0)
Next j%
Dparallelogram(1).data(0).record.data1.index.i(0) = 0
last_conditions.last_cond(0).parallelogram_no = _
     last_conditions.last_cond(0).parallelogram_no + 1
 Dparallelogram(no%).data(0).record.data1.is_removed = True
 End If
 '************************************************************
Case relation_
If ty = 0 Then
Drelation(no%).record_.no_reduce = 255
End If
For i% = 0 To 3
If search_for_relation(Drelation(no%).data(0).data0, i%, n_(i%), 1) = False Then
  remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 3
For j% = n_(i%) + 1 To 2 Step -1
Drelation(j%).data(0).record.data1.index.i(i%) = _
 Drelation(j% - 1).data(0).record.data1.index.i(i%)
Next j%
Drelation(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).relation_no = _
   last_conditions.last_cond(0).relation_no + 1
'last_similar_triangle = last_similar_triangle_for_aid '17
Drelation(no%).data(0).record.data1.is_removed = True
End If
'*****************************************************
Case similar_triangle_
If ty = 0 Then
Dsimilar_triangle(no%).record_.no_reduce = 255
End If
For i% = 0 To 2
If search_for_similar_triangle(Dsimilar_triangle(no%).data(0), i%, n_(i%), 1, 0) = False Then
remove_record = False
End If
Next i%
If remove_record Then
 For i% = 0 To 2
For j% = n_(i%) To 2 Step -1
Dsimilar_triangle(j%).data(0).record.data1.index.i(i%) = _
 Dsimilar_triangle(j% - 1).data(0).record.data1.index.i(i%)
Next j%
Dsimilar_triangle(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).similar_triangle_no = _
     last_conditions.last_cond(0).similar_triangle_no + 1
Dsimilar_triangle(no%).data(0).record.data1.is_removed = True
End If
'*******************************************************
Case total_equal_triangle_
If ty = 0 Then
Dtotal_equal_triangle(no%).record_.no_reduce = 255
End If
For i% = 0 To 2
If search_for_total_equal_triangle(Dtotal_equal_triangle(no%).data(0), i%, n_(i%), 1, 0) = False Then
remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 2
For j% = n_(i%) + 1 To 2 Step -1
Dtotal_equal_triangle(j%).data(0).record.data1.index.i(i%) = _
 Dtotal_equal_triangle(j% - 1).data(0).record.data1.index.i(i%)
Next j%
Dtotal_equal_triangle(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).total_equal_triangle_no = _
     last_conditions.last_cond(0).total_equal_triangle_no + 1
Dtotal_equal_triangle(no%).data(0).record.data1.is_removed = True
End If
'***********************************************************
Case triangle_
For i% = 0 To 2
If search_for_triangle(triangle(no%).data(0), i%, n_(i%), 1) = False Then
 remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 2
For j% = n_(i%) + 1 To 2 Step -1
 triangle(j%).data(0).index.i(i%) = _
  triangle(j% - 1).data(0).index.i(i%)
Next j%
 triangle(1).data(0).index.i(i%) = 0
Next i%
last_conditions.last_cond(0).triangle_no = _
     last_conditions.last_cond(0).triangle_no + 1
End If
'*************************************************************
Case point3_on_line_
If ty = 0 Then
three_point_on_line(no%).record_.no_reduce = 255
End If
For i% = 0 To 2
If search_for_three_point_on_line(three_point_on_line(no%).data(0), 1, i%, n_(i%), 1) = False Then
remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 2
For j% = n_(i%) + 1 To 2 Step -1
three_point_on_line(j%).data(0).record.data1.index.i(i%) = _
 three_point_on_line(j% - 1).data(0).record.data1.index.i(i%)
Next j%
three_point_on_line(0).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).line3_value_no = _
  last_conditions.last_cond(0).line3_value_no + 1
'last_two_line_value = last_two_line_value_for_aid '21
three_point_on_line(no%).data(0).record.data1.is_removed = True
End If
'*********************************************************
Case two_line_value_
If ty = 0 Then
two_line_value(no%).record_.no_reduce = 255
End If
For i% = 0 To 3
If search_for_two_line_value(two_line_value(no%).data(0).data0, i%, n_(i%), 1) = False Then
remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 3
For j% = n_(i%) + 1 To 2 Step -1
two_line_value(j%).data(0).record.data1.index.i(i%) = _
 two_line_value(j% - 1).data(0).record.data1.index.i(i%)
Next j%
two_line_value(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).two_line_value_no = _
   last_conditions.last_cond(0).two_line_value_no + 1
'last_conditions.last_cond(1).line_no3_value = last_conditions.last_cond(1).line_no3_value_for_aid
two_line_value(no%).data(0).record.data1.is_removed = True
End If
'*********************************************************
Case line3_value_
If ty = 0 Then
line3_value(no%).record_.no_reduce = 255
End If
For i% = 0 To 5 '5.7
If search_for_line3_value(line3_value(no%).data(0).data0, i%, n_(i%), 1) = False Then
remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 5
For j% = n_(i%) + 1 To 2 Step -1
line3_value(j%).data(0).record.data1.index.i(i%) = _
 line3_value(j% - 1).data(0).record.data1.index.i(i%)
Next j%
line3_value(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).line3_value_no = _
   last_conditions.last_cond(0).line3_value_no + 1
'last_verti = last_verti_for_aid '24
line3_value(no%).data(0).record.data1.is_removed = True
End If
'*****************************************************
Case verti_
If ty = 0 Then
Dverti(no%).record_.no_reduce = 255
End If
For i% = 0 To 2
If search_for_verti(Dverti(no%).data(0), i%, n_(i%), 1) = False Then '5.7
remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 2
For j% = n_(i%) + 1 To 2 Step -1
Dverti(j%).data(0).record.data1.index.i(i%) = _
 Dverti(j% - 1).data(0).record.data1.index.i(i%)
Next j%
Dverti(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).verti_no = _
   last_conditions.last_cond(0).verti_no + 1
'last_arc_value = last_arc_value_for_aid '25
Dverti(no%).data(0).record.data1.is_removed = True
End If
'**********************************************************
Case arc_value_
If ty = 0 Then
arc_value(no%).record_.no_reduce = 255
End If
For i% = 0 To 1
If search_for_arc_value(arc_value(no%).data(0), 1, n_(i%), 1) = False Then
remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 1
For j% = n_(i%) + 1 To 2 Step -1
arc_value(j%).data(0).record.data1.index.i(0) = _
 arc_value(j% - 1).data(0).record.data1.index.i(0)
Next j%
arc_value(1).data(0).record.data1.index.i(0) = 0
Next i%
last_conditions.last_cond(0).arc_no = _
     last_conditions.last_cond(0).arc_no + 1
     arc_value(no%).data(0).record.data1.is_removed = True
End If
     '*******************************************************
Case equal_arc_
If ty = 0 Then
equal_arc(no%).record_.no_reduce = 255
End If
For i% = 0 To 1
If search_for_equal_arc(equal_arc(no%).data(0), 1, i%, n_(i%), 1) = False Then
remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 1
For j% = n_(i%) + 1 To 2 Step -1
equal_arc(j%).data(0).record.data1.index.i(i%) = _
 equal_arc(j% - 1).data(0).record.data1.index.i(i%)
Next j%
equal_arc(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).equal_arc_no = _
     last_conditions.last_cond(0).equal_arc_no + 1
'last_ratio_of_two_arc = last_ratio_of_two_arc_for_aid '27
'last_angle_less_angle = last_angle_less_angle_for_aid '28
'last_conditions.last_cond(1).line_no_less_line = last_conditions.last_cond(1).line_no_less_line_for_aid '29
'last_conditions.last_cond(1).line_no_less_line2 = last_conditions.last_cond(1).line_no_less_line2_for_aid '30
'last_conditions.last_cond(1).line_no2_less_line2 = last_conditions.last_cond(1).line_no2_less_line2_for_aid '31
'last_angle3_value = last_angle3_value_for_aid '32
equal_arc(no%).data(0).record.data1.is_removed = True
End If
'***************************************************
Case angle3_value_
If ty = 0 Then
angle3_value(no%).record_.no_reduce = 255
End If
For i% = 0 To 5
If search_for_three_angle_value(angle3_value(no%).data(0).data0, _
       i%, n_(i%), 1) = False Then '5.7
       remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 5
For j% = n_(i%) + 1 To 2 Step -1
angle3_value(j%).data(0).record.data1.index.i(i%) = _
 angle3_value(j% - 1).data(0).record.data1.index.i(i%)
Next j%
angle3_value(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).angle3_value_no = last_conditions.last_cond(0).angle3_value_no + 1
'last_conditions.last_cond(1).line_no_value = last_conditions.last_cond(1).line_no_value_for_aid '33
angle3_value(no%).data(0).record.data1.is_removed = True
End If
'*********************************************
Case reduce_angle3_value_
Case line_value_
If ty = 0 Then
line_value(no%).record_.no_reduce = 255
End If
For i% = 0 To 2
If search_for_line_value(line_value(no%).data(0).data0, i%, n_(i%), 1) = False Then
remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 2
For j% = n_(i%) + 1 To 2 Step -1
line_value(j%).data(0).record.data1.index.i(i%) = _
 line_value(j% - 1).data(0).record.data1.index.i(i%)
Next j%
line_value(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).line_value_no = last_conditions.last_cond(0).line_value_no + 1
line_value(no%).data(0).record.data1.is_removed = True
End If
'***********************************************************
'last_tangent_line = last_tangent_line_for_aid '34
'last_tangent_circle = last_tangent_circle_for_aid
'last_equal_area_triangle = last_equal_area_triangle_for_aid '35
'Case equal_area_triangle_
'If ty = 0 Then
'equal_area_triangle(no%).record_.no_reduce = 255
'End If
'For i% = 0 To 1
'Call search_for_equal_area_triangle(equal_area_triangle(no%).data(0), 1, i%, n_(i%), 1)
'For j% = n_(i%) + 1 To 2 Step -1
'equal_area_triangle(j%).data(0).record.data1.index.i(i%) = _
' equal_area_triangle(j% - 1).data(0).record.data1.index.i(i%)
'Next j%
'equal_area_triangle(1).data(0).record.data1.index.i(i%) = 0
'Next i%
'last_conditions.last_cond(0).equal_area_triangle_no = _
'     last_conditions.last_cond(0).equal_area_triangle_no + 1
Case general_string_
If ty = 0 Then
general_string(no%).record_.no_reduce = 255
End If
For i% = 0 To 4
If search_for_general_string(general_string(no%).data(0), i%, n_(i%), 1) = False Then
remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 4
For j% = n_(i%) + 1 To 2 Step -1
general_string(j%).data(0).record.data1.index.i(i%) = _
 general_string(j% - 1).data(0).record.data1.index.i(i%)
Next j%
general_string(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).general_string_no = _
     last_conditions.last_cond(0).general_string_no + 1
general_string(no%).data(0).record.data1.is_removed = True
End If
'*************************************************************
'last_general_angle_string = last_general_angle_string_for_aid '37
'last_conditions.last_cond(1).equal_side_tixing_no = last_conditions.last_cond(1).equal_side_tixing_no_for_aid '38
'last_conditions.last_cond(1).Epolygon_no = last_conditions.last_cond(1).Epolygon_no_for_aid '39
Case epolygon_
If ty = 0 Then
epolygon(no%).record_.no_reduce = 255
End If
If search_for_epolygon(epolygon(no%).data(0), 1, n_(0), 1) Then
remove_record = False
For j% = n_(0) + 1 To 2 Step -1
epolygon(j%).data(0).record.data1.index.i(0) = _
 epolygon(j% - 1).data(0).record.data1.index.i(0)
Next j%
epolygon(1).data(0).record.data1.index.i(0) = 0
last_conditions.last_cond(0).epolygon_no = _
     last_conditions.last_cond(0).epolygon_no + 1
epolygon(no%).data(0).record.data1.is_removed = True
End If
'************************************************************
'last_conditions.last_cond(1).tixing_no = last_conditions.last_cond(1).tixing_no_for_aid '40
'last_conditions.last_cond(1).rhombus_no = last_conditions.last_cond(1).rhombus_no_for_aid '41
'last_conditions.last_cond(1).last_long_squre_no = last_conditions.last_cond(1).last_long_squre_no_for_aid '42
'last_area_of_triangle = last_area_of_triangle_for_aid '43
Case area_of_element_
If ty = 0 Then
area_of_element(no%).record_.no_reduce = 255
End If
If search_for_area_element(area_of_element(no%).data(0), 1, n_(0), 1) Then
For j% = n_(0) + 1 To 2 Step -1
area_of_element(j%).data(0).record.data1.index.i(0) = _
  area_of_element(j% - 1).data(0).record.data1.index.i(0)
Next j%
area_of_element(1).data(0).record.data1.index.i(0) = 0
last_conditions.last_cond(0).area_of_element_no = _
  last_conditions.last_cond(0).area_of_element_no + 1
area_of_element(no%).data(0).record.data1.is_removed = True
End If
'*****************************************************************
Case sides_length_of_triangle_
If ty = 0 Then
Sides_length_of_triangle(no%).record_.no_reduce = 255
End If
If search_for_sides_length_of_triangle(Sides_length_of_triangle(no%).data(0), 1, n_(0), 1) Then
For j% = n_(0) + 1 To 2 Step -1
Sides_length_of_triangle(j%).data(0).record.data1.index.i(0) = _
  Sides_length_of_triangle(j% - 1).data(0).record.data1.index.i(0)
Next j%
Sides_length_of_triangle(1).data(0).record.data1.index.i(0) = 0
last_conditions.last_cond(0).sides_length_of_triangle_no = _
     last_conditions.last_cond(0).sides_length_of_triangle_no + 1
Sides_length_of_triangle(no%).data(0).record.data1.is_removed = True
End If
'*********************************************************************
Case verti_mid_line_
If ty = 0 Then
verti_mid_line(no%).record_.no_reduce = 255
End If
For i% = 0 To 1
If search_for_verti_mid_line(verti_mid_line(no%).data(0).data0, n_(i%), 1, 1) = False Then
remove_record = False
End If
Next i%
If remove_record Then
For i% = 0 To 1
For j% = n_(i%) + 1 To 2 Step -1
verti_mid_line(j%).data(0).record.data1.index.i(i%) = _
 verti_mid_line(j% - 1).data(0).record.data1.index.i(i%)
Next j%
verti_mid_line(1).data(0).record.data1.index.i(i%) = 0
Next i%
last_conditions.last_cond(0).verti_mid_line_no = _
     last_conditions.last_cond(0).verti_mid_line_no + 1
verti_mid_line(no%).data(0).record.data1.is_removed = True
End If
'last_sides_length_of_circle = last_sides_length_of_circle_for_aid '48
'last_verti_mid_line = last_verti_mid_line_for_aid '49
'last_conditions.last_cond(1).point_no_in_mid_verti_line = last_conditions.last_cond(1).point_no_in_mid_verti_line_for_aid
'last_squ_sum = last_squ_sum_for_aid '50
End Select
remove_record_error:
End Function


Public Function simple_three_angle_value(ByVal replace_no%, tv$, ByVal no%, _
  re As record_data_type, ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, tn_%, l%
Dim n(2) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim t_a_v() As angle3_value_data0_type
Dim last_tn%
Dim tA(1) As Integer
Dim ty As Byte
Dim t_A As angle3_value_data0_type
Dim temp_record As total_record_type
Dim temp_record1 As total_record_type
i% = angle(replace_no%).data(0).total_no
j% = angle(no%).data(0).total_no
k% = T_angle(i%).data(0).is_used_no '
tn_% = T_angle(j%).data(0).is_used_no
If (k% - tn_%) Mod 2 = 0 Then
tA(1) = T_angle(i%).data(0).angle_no(k%).no
tA(0) = T_angle(j%).data(0).angle_no(k%).no
Else
tA(1) = T_angle(i%).data(0).angle_no(k%).no
tA(0) = T_angle(j%).data(0).angle_no(tn_%).no
If tv <> "" Then
  tv$ = minus_string("180", tv$, True, False)
End If
ty = 1
End If
last_tn% = 0
For i% = 0 To 2
n(0) = i%
n(1) = (i% + 1) Mod 3
n(2) = (i% + 2) Mod 3
For k% = 0 To 1
 t_A.angle(n(0)) = tA(k%)
 t_A.angle(n(1)) = -1
Call search_for_three_angle_value(t_A, n(0), n_(0), 1)  '5.7
 t_A.angle(n(1)) = 30000
Call search_for_three_angle_value(t_A, n(0), n_(1), 1)
For j% = n_(0) + 1 To n_(1)
tn_% = angle3_value(j%).data(0).record.data1.index.i(n(0))
If tn_% > 0 Then
 For l% = 1 To last_tn%
   If tn_% = tn(l%) Then
    If ty = 0 Then
     If tv$ <> "" Then
        t_a_v(l%).value = minus_string(t_a_v(l%).value, _
            time_string("180", t_a_v(l%).para(i%), False, False), True, False)
        t_a_v(l%).angle(i%) = 0
     Else
        t_a_v(l%).angle(i%) = tA(k%)
     End If
    Else
     If tv$ <> "" Then
        t_a_v(l%).value = minus_string(t_a_v(l%).value, _
            time_string("180", t_a_v(l%).para(i%), False, False), True, False)
        t_a_v(l%).angle(i%) = 0
     Else
        t_a_v(l%).angle(i%) = tA(k%)
        t_a_v(l%).para(i%) = time_string("-1", t_a_v(l%).para(i%), True, False)
        t_a_v(l%).value = add_string(t_a_v(l%).value, _
            time_string("180", t_a_v(l%).para(i%), False, False), True, False)
     End If
    End If
    GoTo simple_three_angle_value_next
   End If
 Next l%
 last_tn% = last_tn% + 1
 ReDim Preserve tn(last_tn%) As Integer
 ReDim Preserve t_a_v(last_tn%) As angle3_value_data0_type
 t_a_v(last_tn%) = angle3_value(tn_%).data(0).data0
 If k% = 1 Then
  If ty = 0 Then
     t_a_v(last_tn%).angle(i%) = tA(k%)
  Else
     t_a_v(last_tn%).angle(i%) = tA(k%)
     t_a_v(last_tn%).para(i%) = time_string("-1", t_a_v(last_tn%).para(i%), True, False)
     t_a_v(last_tn%).value = add_string(t_a_v(last_tn%).value, _
          time_string("180", t_a_v(last_tn%).para(i%), False, False), True, False)
  End If
 End If
 tn(last_tn%) = tn_%
End If
simple_three_angle_value_next:
Next j%
Next k%
Next i%
For j% = 1 To last_tn%
tn_% = tn(j%)
temp_record.record_data = re
Call add_conditions_to_record(angle3_value_, tn_%, 0, 0, temp_record.record_data.data0.condition_data)
Call remove_record(angle3_value_, tn_%, 0)
 simple_three_angle_value = set_three_angle_value(t_a_v(j%).angle(0), t_a_v(j%).angle(1), _
   t_a_v(j%).angle(2), t_a_v(j%).para(0), t_a_v(j%).para(1), t_a_v(j%).para(2), _
     t_a_v(j%).value, 0, temp_record, 0, 0, 0, 0, 0, 0, False)
 If simple_three_angle_value > 1 Then
    Exit Function
 End If
Next j%
End Function


Public Sub simple_record_by_replace(re As record_data_type, _
               ByVal replace_ty As Byte, ByVal con_ty As Byte, ByVal replace_no%, ByVal n%)
Dim i%
 For i% = 1 To re.data0.condition_data.condition_no
  If re.data0.condition_data.condition(i%).ty = con_ty Then
    If re.data0.condition_data.condition(i%).no = n% Then
      re.data0.condition_data.condition(i%).ty = replace_ty
        re.data0.condition_data.condition(i%).no = replace_no%
         Call set_level(re.data0.condition_data)
          Exit Sub
    End If
  End If
 Next i%
End Sub
Public Sub add_record(ByVal record_type As Byte, ByVal no%, is_remove As Boolean)
Dim i%, j%, k%
Dim n_(8) As Integer
Select Case record_type
Case item0_
For i% = 0 To 3
Call search_for_item0(item0(no%).data(0), i%, n_(i%), 1)
For j% = 1 To n_(i%) - 1
item0(j%).data(0).index(i%) = _
         item0(j% + 1).data(0).index(i%)
Next j%
item0(n_(i%)).data(0).index(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).item0_no = last_conditions.last_cond(0).item0_no - 1
End If
'********************************************************
Case angle_
For i% = 0 To 1
Call search_for_angle(angle(no%).data(0), n_(i%), i%, 1)
For j% = 1 To n_(i%) - 1
angle(j%).data(0).index(i%) = _
         angle(j% + 1).data(0).index(i%)
Next j%
angle(n_(i%)).data(0).index(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).angle_no = last_conditions.last_cond(0).angle_no - 1
End If
'***********************************************************************************
Case line_from_two_point_
Call search_for_two_point_line(Dtwo_point_line(no%).poi(0), _
       Dtwo_point_line(no%).poi(1), n_(0), 1)
For j% = 1 To n_(0) - 1
Dtwo_point_line(j%).data(0).index = _
 Dtwo_point_line(j% + 1).data(0).index
Next j%
Dtwo_point_line(n_(0)).data(0).index = no%
If is_remove Then
last_conditions.last_cond(1).line_from_two_point_no = _
  last_conditions.last_cond(1).line_from_two_point_no - 1
End If
'*************************************************************
Case dpoint_pair_
For i% = 0 To 6
Call search_for_point_pair(Ddpoint_pair(no%).data(0).data0, i%, n_(i%), 1)
For j% = 1 To n_(i%) - 1
Ddpoint_pair(j%).data(0).record.data1.index.i(i%) = _
 Ddpoint_pair(j% + 1).data(0).record.data1.index.i(i%)
Next j%
Ddpoint_pair(n_(i%)).data(0).record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).dpoint_pair_no = last_conditions.last_cond(0).dpoint_pair_no - 1
End If
'**********************************************************************
Case area_relation_
For i% = 0 To 1
Call search_for_area_relation(Darea_relation(no%).data(0), i%, n_(i%), 1)
For j% = 1 To n_(i%) - 1
Darea_relation(j%).data(0).record.data1.index.i(i%) = _
Darea_relation(j% + 1).data(0).record.data1.index.i(i%)
Next j%
Darea_relation(n_(i%)).data(0).record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).area_relation_no = last_conditions.last_cond(0).area_relation_no - 1
End If
'**********************************************************************
''last_Dpoint = last_dpoint_for_aid '5
'last_mid_point_line = last_mid_point_line_for_aid '7
'last_eline = last_eline_for_aid '8
Case eline_
For i% = 0 To 3
Call search_for_eline(Deline(no%).data(0).data0, i%, n_(i%), 1)
For j% = 1 To n_(i%) - 1
Deline(j%).data(0).record.data1.index.i(i%) = _
 Deline(j% + 1).data(0).record.data1.index.i(i%)
Next j%
Deline(n_(i%)).data(0).record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).eline_no = last_conditions.last_cond(0).eline_no - 1
End If
'last_Four_point_on_circle = last_four_point_on_circle_for_aid '9
'****************************************************
Case point4_on_circle_
For i% = 0 To 3
Call search_for_four_point_on_circle(four_point_on_circle(no%).data(0), 1, i%, n_(i%), 1)
For j% = 1 To n_(i%) - 1
four_point_on_circle(j%).data(0).record.data1.index.i(i%) = _
 four_point_on_circle(j% + 1).data(0).record.data1.index.i(i%)
Next j%
four_point_on_circle(n_(i%)).data(0).record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).four_point_on_circle_no = last_conditions.last_cond(0).four_point_on_circle_no - 1
End If
'**************************************************************
Case midpoint_
For i% = 0 To 2
Call search_for_mid_point(Dmid_point(no%).data(0).data0, i%, n_(i%), 1) '5.7
For j% = 1 To n_(i%) - 1
Dmid_point(j%).data(0).record.data1.index.i(i%) = _
 Dmid_point(j% + 1).data(0).record.data1.index.i(i%)
Next j%
Dmid_point(n_(i%)).data(0).record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).new_point_no = last_conditions.last_cond(0).new_point_no - 1
End If
'last_paral = last_paral_for_aid '13
'**************************************************************
Case paral_
For i% = 0 To 2
Call search_for_paral(Dparal(no%).data(0).data0, i%, n_(i%), 1)
For j% = 1 To n_(i%) - 1
Dparal(j%).data(0).data0.record.data1.index.i(i%) = _
 Dparal(j% + 1).data(0).data0.record.data1.index.i(i%)
Next j%
Dparal(n_(i%)).data(0).data0.record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).dpoint_pair_no = last_conditions.last_cond(0).dpoint_pair_no - 1
End If
'last_parallelogram = last_parallelogram_for_aid '14
'***********************************************************
Case parallelogram_
'For i% = 0 To 3
Call search_for_parallelogram(Dparallelogram(no%).data(0).polygon4_no, n_(i%), 1)
For j% = 0 To n_(i%) - 1
Dparallelogram(j%).data(0).record.data1.index.i(0) = _
 Dparallelogram(j% + 1).data(0).record.data1.index.i(0)
Next j%
Dparallelogram(n_(i%)).data(0).record.data1.index.i(0) = no%
If is_remove Then
last_conditions.last_cond(0).parallelogram_no = last_conditions.last_cond(0).parallelogram_no - 1
End If
'***********************************************************
'last_conditions.last_cond(1).point_no = last_aid_point '15
'Last_relation = last_relation_for_aid '16
Case relation_
For i% = 0 To 3
Call search_for_relation(Drelation(no%).data(0).data0, i%, n_(i%), 1)
For j% = 1 To n_(i%)
Drelation(j%).data(0).record.data1.index.i(i%) = _
 Drelation(j% + 1).data(0).record.data1.index.i(i%)
Next j%
Drelation(n_(i%)).data(0).record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).relation_no = last_conditions.last_cond(0).relation_no - 1
End If
'********************************************************
'last_similar_triangle = last_similar_triangle_for_aid '17
Case similar_triangle_
For i% = 0 To 2
Call search_for_similar_triangle(Dsimilar_triangle(no%).data(0), i%, n_(i%), 1, 0)
For j% = 1 To n_(i%) - 1
Dsimilar_triangle(j%).data(0).record.data1.index.i(i%) = _
 Dsimilar_triangle(j% + 1).data(0).record.data1.index.i(i%)
Next j%
Dsimilar_triangle(n_(i%)).data(0).record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).similar_triangle_no = last_conditions.last_cond(0).similar_triangle_no - 1
End If
'*********************************************************
'Last_total_equal_triangle = last_total_equal_triangle_for_aid '18
Case total_equal_triangle_
For i% = 0 To 2
Call search_for_total_equal_triangle(Dtotal_equal_triangle(no%).data(0), i%, n_(i%), 1, 0)
For j% = 1 To n_(i%) - 1
Dtotal_equal_triangle(j%).data(0).record.data1.index.i(i%) = _
 Dtotal_equal_triangle(j% + 1).data(0).record.data1.index.i(i%)
Next j%
Dtotal_equal_triangle(n_(i%)).data(0).record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).total_equal_triangle_no = last_conditions.last_cond(0).total_equal_triangle_no - 1
End If
'*********************************************************
'last_triangle = last_triangle_for_aid '19
Case triangle_
For i% = 0 To 2
Call search_for_triangle(triangle(no%).data(0), i%, n_(i%), 1)
For j% = 1 To n_(i%) - 1
 triangle(j%).data(0).index.i(i%) = _
  triangle(j% + 1).data(0).index.i(i%)
Next j%
 triangle(n_(i%)).data(0).index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).triangle_no = last_conditions.last_cond(0).triangle_no - 1
End If
'**********************************************************
'last_Three_point_on_line = last_three_point_on_line_for_aid '20
Case point3_on_line_
For i% = 0 To 2
Call search_for_three_point_on_line(three_point_on_line(no%).data(0), 1, i%, n_(i%), 1)
For j% = 1 To n_(i%) - 1
three_point_on_line(j%).data(0).record.data1.index.i(i%) = _
 three_point_on_line(j% + 1).data(0).record.data1.index.i(i%)
Next j%
Next i%
three_point_on_line(n_(i%)).data(0).record.data1.index.i(i%) = no%
If is_remove Then
last_conditions.last_cond(0).three_point_on_line_no = last_conditions.last_cond(0).three_point_on_line_no - 1
End If
'**************************************************************
'last_two_line_value = last_two_line_value_for_aid '21
Case two_line_value_
For i% = 0 To 3
Call search_for_two_line_value(two_line_value(no%).data(0).data0, i%, n_(i%), 1)
For j% = 1 To n_(i%) - 1
two_line_value(j%).data(0).record.data1.index.i(i%) = _
 two_line_value(j% + 1).data(0).record.data1.index.i(i%)
Next j%
two_line_value(n_(i%)).data(0).record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).two_line_value_no = last_conditions.last_cond(0).two_line_value_no - 1
End If
'***************************************************************
'last_conditions.last_cond(1).line_no3_value = last_conditions.last_cond(1).line_no3_value_for_aid
Case line3_value_
For i% = 0 To 5 '5.7
Call search_for_line3_value(line3_value(no%).data(0).data0, i%, n_(i%), 1)
For j% = 1 To n_(i%) - 1
line3_value(j%).data(0).record.data1.index.i(i%) = _
 line3_value(j% + 1).data(0).record.data1.index.i(i%)
Next j%
line3_value(n_(i%)).data(0).record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).line3_value_no = last_conditions.last_cond(0).line3_value_no - 1
End If
'************************************************************
'last_verti = last_verti_for_aid '24
Case verti_
For i% = 0 To 2
Call search_for_verti(Dverti(no%).data(0), i%, n_(i%), 1) '5.7
For j% = 1 To n_(i%) - 1
Dverti(j%).data(0).record.data1.index.i(i%) = _
 Dverti(j% + 1).data(0).record.data1.index.i(i%)
Next j%
Dverti(n_(i%)).data(0).record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).verti_no = last_conditions.last_cond(0).verti_no - 1
End If
'*******************************************************
'last_arc_value = last_arc_value_for_aid '25
Case arc_value_
For i% = 0 To 1
Call search_for_arc_value(arc_value(no%).data(0), 1, n_(i%), 1)
For j% = 1 To n_(0)
arc_value(j%).data(0).record.data1.index.i(i%) = _
 arc_value(j% + 1).data(0).record.data1.index.i(i%)
Next j%
arc_value(n_(0) + 1).data(0).record.data1.index.i(i%) = no%
Next i%
'last_equal_arc = last_equal_arc_for_aid '26
If is_remove Then
last_conditions.last_cond(0).arc_value_no = last_conditions.last_cond(0).arc_value_no - 1
End If
'***************************************************************
Case equal_arc_
For i% = 0 To 1
If search_for_equal_arc(equal_arc(no%).data(0), 1, i%, n_(i%), 1) = False Then
For j% = 1 To n_(i%) - 1
equal_arc(j%).data(0).record.data1.index.i(i%) = _
 equal_arc(j% + 1).data(0).record.data1.index.i(i%)
Next j%
equal_arc(n_(i%)).data(0).record.data1.index.i(i%) = no%
End If
Next i%
If is_remove Then
last_conditions.last_cond(0).equal_arc_no = last_conditions.last_cond(0).equal_arc_no - 1
End If
'**********************************************************
'last_ratio_of_two_arc = last_ratio_of_two_arc_for_aid '27
'last_angle_less_angle = last_angle_less_angle_for_aid '28
'last_conditions.last_cond(1).line_no_less_line = last_conditions.last_cond(1).line_no_less_line_for_aid '29
'last_conditions.last_cond(1).line_no_less_line2 = last_conditions.last_cond(1).line_no_less_line2_for_aid '30
'last_conditions.last_cond(1).line_no2_less_line2 = last_conditions.last_cond(1).line_no2_less_line2_for_aid '31
'last_angle3_value = last_angle3_value_for_aid '32
Case angle3_value_
For i% = 0 To 5
Call search_for_three_angle_value(angle3_value(no%).data(0).data0, _
        i%, n_(i%), 1) '5.7
For j% = 1 To n_(i%) - 1
angle3_value(j%).data(0).record.data1.index.i(i%) = _
 angle3_value(j% + 1).data(0).record.data1.index.i(i%)
Next j%
angle3_value(n_(i%)).data(0).record.data1.index.i(i%) = 0
Next i%
If is_remove Then
last_conditions.last_cond(0).line3_value_no = last_conditions.last_cond(0).line3_value_no - 1
End If
'***********************************************************
'last_conditions.last_cond(1).line_no_value = last_conditions.last_cond(1).line_no_value_for_aid '33
Case line_value_
For i% = 0 To 2
Call search_for_line_value(line_value(no%).data(0).data0, i%, n_(i%), 1)
For j% = 1 To n_(i%) - 1
line_value(j%).data(0).record.data1.index.i(i%) = _
 line_value(j% + 1).data(0).record.data1.index.i(i%)
Next j%
line_value(n_(i%)).data(0).record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).line_value_no = last_conditions.last_cond(0).line_value_no - 1
End If
'***************************************************
'last_tangent_line = last_tangent_line_for_aid '34
'last_tangent_circle = last_tangent_circle_for_aid
'last_equal_area_triangle = last_equal_area_triangle_for_aid '35
'Case equal_area_triangle_
'For i% = 0 To 1
'Call search_for_equal_area_triangle(equal_area_triangle(no%).data(0), 1, i%, n_(i%), 1)
'For j% = 1 To n_(i%) - 1
'equal_area_triangle(j%).data(0).record.data1.index.i(i%) = _
' equal_area_triangle(j% + 1).data(0).record.data1.index.i(i%)
'Next j%
'equal_area_triangle(n_(i%)).data(0).record.data1.index.i(i%) = no%
'Next i%
'last_general_string = last_general_string_for_aid '36
Case general_string_
For i% = 0 To 4
Call search_for_general_string(general_string(no%).data(0), i%, n_(i%), 1)
For j% = 1 To n_(i%) - 1
general_string(j%).data(0).record.data1.index.i(i%) = _
 general_string(j% + 1).data(0).record.data1.index.i(i%)
Next j%
general_string(n_(i%)).data(0).record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).general_string_no = last_conditions.last_cond(0).general_string_no - 1
End If
'**********************************************
'last_general_angle_string = last_general_angle_string_for_aid '37
'last_conditions.last_cond(1).equal_side_tixing_no = last_conditions.last_cond(1).equal_side_tixing_no_for_aid '38
'last_conditions.last_cond(1).Epolygon_no = last_conditions.last_cond(1).Epolygon_no_for_aid '39
Case epolygon_
If search_for_epolygon(epolygon(no%).data(0), 1, n_(0), 1) = False Then
For j% = 1 To n_(0) - 1
epolygon(j%).data(0).record.data1.index.i(0) = _
 epolygon(j% + 1).data(0).record.data1.index.i(0)
Next j%
epolygon(n_(0)).data(0).record.data1.index.i(0) = no%
End If
If is_remove Then
last_conditions.last_cond(0).epolygon_no = last_conditions.last_cond(0).epolygon_no - 1
End If
'************************************************
'last_conditions.last_cond(1).tixing_no = last_conditions.last_cond(1).tixing_no_for_aid '40
'last_conditions.last_cond(1).rhombus_no = last_conditions.last_cond(1).rhombus_no_for_aid '41
'last_conditions.last_cond(1).last_long_squre_no = last_conditions.last_cond(1).last_long_squre_no_for_aid '42
'last_area_of_triangle = last_area_of_triangle_for_aid '43
Case area_of_element_
Call search_for_area_element(area_of_element(no%).data(0), 1, n_(0), 1)
For j% = 1 To n_(0) - 1
area_of_element(j%).data(0).record.data1.index.i(0) = _
  area_of_element(j% + 1).data(0).record.data1.index.i(0)
Next j%
area_of_element(n_(0) + 1).data(0).record.data1.index.i(0) = no%
If is_remove Then
last_conditions.last_cond(0).area_of_element_no = last_conditions.last_cond(0).area_of_element_no - 1
End If
'last_area_of_circle = last_area_of_circle_for_aid '44
'last_area_of_polygon = last_area_of_polygon_for_aid '45
'********************************************
Case sides_length_of_triangle_
Call search_for_sides_length_of_triangle(Sides_length_of_triangle(no%).data(0), 1, n_(0), 1)
For j% = 1 To n_(0) - 1
Sides_length_of_triangle(j%).data(0).record.data1.index.i(0) = _
  Sides_length_of_triangle(j% + 1).data(0).record.data1.index.i(0)
Next j%
Sides_length_of_triangle(n_(0)).data(0).record.data1.index.i(0) = no%
If is_remove Then
last_conditions.last_cond(0).sides_length_of_triangle_no = _
                last_conditions.last_cond(0).sides_length_of_triangle_no - 1
End If
Case verti_mid_line_
For i% = 0 To 1
Call search_for_verti_mid_line(verti_mid_line(no%).data(0).data0, n_(i%), 1, 1) '5.7
For j% = 1 To n_(i%) - 1
verti_mid_line(j%).data(0).record.data1.index.i(i%) = _
  verti_mid_line(j% + 1).data(0).record.data1.index.i(i%)
Next j%
verti_mid_line(n_(0)).data(0).record.data1.index.i(i%) = no%
Next i%
If is_remove Then
last_conditions.last_cond(0).verti_mid_line_no = _
                last_conditions.last_cond(0).verti_mid_line_no
End If
'last_sides_length_of_circle = last_sides_length_of_circle_for_aid '48
'last_verti_mid_line = last_verti_mid_line_for_aid '49
'last_conditions.last_cond(1).point_no_in_mid_verti_line = last_conditions.last_cond(1).point_no_in_mid_verti_line_for_aid
'last_squ_sum = last_squ_sum_for_aid '50
End Select

End Sub
Public Function set_total_equal_triangle_from_simple_angle(ByVal A%, ByVal new_p%, re As total_record_type) As Byte
Dim A3_v As angle3_value_data0_type
Dim i%, j%, k%, no%, l%, p%, tA%
Dim last_tn%, last_tn1%, last_tn2%, last_tn3% ', last_tn4%
Dim n(2) As Integer
Dim m(2) As Integer
'Dim m_(0) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim tn1() As Integer
Dim tn2() As Integer
Dim tn3() As Integer
'Dim tn4() As Integer
Dim tA1(2) As Integer
Dim tA2(2) As Integer
Dim s1(2) As String
Dim S2(2) As String
Dim v(1) As String
Dim ty As Byte
Dim t_A As angle3_value_data0_type
Dim temp_record As total_record_type
temp_record = re
If angle(A%).data(0).value <> "" Then
For i% = last_conditions.last_cond(0).angle_no + 1 To last_conditions.last_cond(1).angle_no
 no% = angle(i%).data(0).index(0)
  If no% < A% Then
   If is_point_in_line3(new_p%, m_lin(angle(no%).data(0).line_no(0)).data(0).data0, 0) Or _
       is_point_in_line3(new_p%, m_lin(angle(no%).data(0).line_no(1)).data(0).data0, 0) Then
  temp_record = re
   Call add_conditions_to_record(angle3_value_, angle(no%).data(0).value_no, angle(A%).data(0).value_no, _
        0, temp_record.record_data.data0.condition_data)
  If angle(no%).data(0).value = angle(A%).data(0).value Then
  set_total_equal_triangle_from_simple_angle = set_total_equal_triangle_from_eangle(A%, no%, _
       temp_record, 0, 0, 0, 0, 0, 0, 1)
  End If
 End If
 End If
Next i%
Else
For j% = 0 To 1
 m(0) = j%
  m(1) = (j% + 1) Mod 3
   m(2) = (j% + 2) Mod 3
t_A.angle(m(0)) = A%
t_A.angle(m(1)) = -1
Call search_for_three_angle_value(t_A, j%, n_(0), 1)   '5.7
t_A.angle(m(1)) = 30000
Call search_for_three_angle_value(t_A, j%, n_(1), 1)   '5.7
 m(1) = (m(0) + 1) Mod 2
For k% = n_(0) + 1 To n_(1)
no% = angle3_value(k%).data(0).record.data1.index.i(j%)
If angle3_value(no%).data(0).data0.type = eangle_ Then
 If angle3_value(no%).data(0).data0.angle(m(1)) < A% Then
  If is_point_in_line3(new_p%, m_lin(angle(angle3_value(no%).data(0).data0.angle(m(1))).data(0).line_no(0)).data(0).data0, 0) Or _
       is_point_in_line3(new_p%, m_lin(angle(angle3_value(no%).data(0).data0.angle(m(1))).data(0).line_no(1)).data(0).data0, 0) Then
  last_tn% = last_tn% + 1
   ReDim Preserve tn(last_tn%) As Integer
    tn(last_tn%) = no%
  End If
 End If
End If
Next k%
'******************************************
For i% = 1 To last_tn%
 no% = tn(i%)
 temp_record = re
    Call add_conditions_to_record(angle3_value_, no%, 0, _
        0, temp_record.record_data.data0.condition_data)
  set_total_equal_triangle_from_simple_angle = set_total_equal_triangle_from_eangle(angle3_value(no%).data(0).data0.angle(0), _
    angle3_value(no%).data(0).data0.angle(1), temp_record, new_p%, 0, 0, 0, 0, 0, 1)
  If set_total_equal_triangle_from_simple_angle > 0 Then
   Exit Function
  End If
Next i%
Next j%
End If
End Function

Public Function set_total_equal_triangle_from_combine_two_line(ByVal replace_l%, ByVal new_p%, re As total_record_type) As Byte
Dim i%
Dim no%
For i% = last_conditions.last_cond(0).angle_no + 1 To last_conditions.last_cond(1).angle_no
no% = angle(i%).data(0).index(0)
If angle(no%).data(0).line_no(0) = replace_l% Or angle(no%).data(0).line_no(1) = replace_l% Then
 set_total_equal_triangle_from_combine_two_line = _
  set_total_equal_triangle_from_simple_angle(no%, new_p%, re)
   If set_total_equal_triangle_from_combine_two_line > 1 Then
    Exit Function
   End If
End If
Next i%
End Function
Public Sub simple_data_for_add_point_to_line(ByVal p%, ByVal l%)
Dim i%, j%, n%
If is_point_in_line3(p%, m_lin(l%).data(0).data0, n%) Then
For i% = 1 To last_conditions.last_cond(1).line_value_no
     If line_value(i%).data(0).data0.line_no = l% Then
      Call is_point_in_line3(line_value(i%).data(0).data0.poi(0), _
               m_lin(l%).data(0).data0, line_value(i%).data(0).data0.n(0))
      Call is_point_in_line3(line_value(i%).data(0).data0.poi(1), _
               m_lin(l%).data(0).data0, line_value(i%).data(0).data0.n(1))
     End If
Next i%
For i% = 1 To last_conditions.last_cond(1).two_line_value_no
  For j% = 0 To 2
     If two_line_value(i%).data(0).data0.line_no(j%) = l% Then
      Call is_point_in_line3(two_line_value(i%).data(0).data0.poi(2 * j), _
               m_lin(l%).data(0).data0, two_line_value(i%).data(0).data0.n(2 * j%))
      Call is_point_in_line3(two_line_value(i%).data(0).data0.poi(2 * j% + 1), _
               m_lin(l%).data(0).data0, two_line_value(i%).data(0).data0.n(2 * j% + 1))
     End If
  Next j%
Next i%
For i% = 1 To last_conditions.last_cond(1).line3_value_no
  For j% = 0 To 2
     If line3_value(i%).data(0).data0.line_no(j%) = l% Then
      Call is_point_in_line3(line3_value(i%).data(0).data0.poi(2 * j), _
               m_lin(l%).data(0).data0, line3_value(i%).data(0).data0.n(2 * j%))
      Call is_point_in_line3(line3_value(i%).data(0).data0.poi(2 * j% + 1), _
               m_lin(l%).data(0).data0, line3_value(i%).data(0).data0.n(2 * j% + 1))
     End If
  Next j%
Next i%
For i% = 1 To last_conditions.last_cond(1).mid_point_no
 If Dmid_point(i%).data(0).data0.line_no = l% Then
      Call is_point_in_line3(Dmid_point(i%).data(0).data0.poi(0), _
               m_lin(l%).data(0).data0, Dmid_point(i%).data(0).data0.n(0))
      Call is_point_in_line3(Dmid_point(i%).data(0).data0.poi(1), _
               m_lin(l%).data(0).data0, Dmid_point(i%).data(0).data0.n(1))
      Call is_point_in_line3(Dmid_point(i%).data(0).data0.poi(2), _
               m_lin(l%).data(0).data0, Dmid_point(i%).data(0).data0.n(2))
 End If
Next i%
For i% = 1 To last_conditions.last_cond(1).eline_no
  For j% = 0 To 1
     If Deline(i%).data(0).data0.line_no(j%) = l% Then
      Call is_point_in_line3(Deline(i%).data(0).data0.poi(2 * j), _
               m_lin(l%).data(0).data0, Deline(i%).data(0).data0.n(2 * j%))
      Call is_point_in_line3(Deline(i%).data(0).data0.poi(2 * j% + 1), _
               m_lin(l%).data(0).data0, Deline(i%).data(0).data0.n(2 * j% + 1))
     End If
  Next j%
Next i%
For i% = 1 To last_conditions.last_cond(1).relation_no
  For j% = 0 To 2
     If Drelation(i%).data(0).data0.line_no(j%) = l% Then
      Call is_point_in_line3(Drelation(i%).data(0).data0.poi(2 * j), _
               m_lin(l%).data(0).data0, Drelation(i%).data(0).data0.n(2 * j%))
      Call is_point_in_line3(Drelation(i%).data(0).data0.poi(2 * j% + 1), _
               m_lin(l%).data(0).data0, Drelation(i%).data(0).data0.n(2 * j% + 1))
     End If
  Next j%
Next i%
For i% = 1 To last_conditions.last_cond(1).dpoint_pair_no
  For j% = 0 To 5
     If Ddpoint_pair(i%).data(0).data0.line_no(j%) = l% Then
      Call is_point_in_line3(Ddpoint_pair(i%).data(0).data0.poi(2 * j), _
               m_lin(l%).data(0).data0, Ddpoint_pair(i%).data(0).data0.n(2 * j%))
      Call is_point_in_line3(Ddpoint_pair(i%).data(0).data0.poi(2 * j% + 1), _
               m_lin(l%).data(0).data0, Ddpoint_pair(i%).data(0).data0.n(2 * j% + 1))
     End If
  Next j%
Next i%
For i% = 1 To last_conditions.last_cond(1).item0_no
  For j% = 0 To 2
     If item0(i%).data(0).line_no(j%) = l% Then
      Call is_point_in_line3(item0(i%).data(0).poi(2 * j), _
               m_lin(l%).data(0).data0, item0(i%).data(0).n(2 * j%))
      Call is_point_in_line3(item0(i%).data(0).poi(2 * j% + 1), _
               m_lin(l%).data(0).data0, item0(i%).data(0).n(2 * j% + 1))
     End If
  Next j%
Next i%
End If
End Sub
Public Function simple_dbase_for_angle0(ByVal no%, ByVal replace_no%, ByVal ts$, re As total_record_type) As Byte
Dim n1%, n2%
Dim temp_record As total_record_type
temp_record = re
If simple_dbase_for_angle0 > 1 Then
   Exit Function
End If
If ts$ = "" Then
n1% = T_angle(angle(no%).data(0).total_no).data(0).is_used_no
n2% = T_angle(angle(replace_no%).data(0).total_no).data(0).is_used_no
angle(no%).data(0).other_no = replace_no%
       If (Abs(n1% - n2%)) Mod 2 = 0 Then
           temp_record = re
           simple_dbase_for_angle0 = simple_dbase_for_angle( _
             T_angle(angle(no%).data(0).total_no).data(0).angle_no(n1%).no, _
              T_angle(angle(replace_no%).data(0).total_no).data(0).angle_no(n2%).no, 0, re)
       Else
            temp_record = re
            simple_dbase_for_angle0 = simple_dbase_for_angle( _
             T_angle(angle(no%).data(0).total_no).data(0).angle_no(n1%).no, _
              T_angle(angle(replace_no%).data(0).total_no).data(0).angle_no(n2%).no, 1, re)
       End If
        If simple_dbase_for_angle0 > 1 Then
          Exit Function
        End If
Else
 Call remove_record(angle_, no%, 0)
End If
End Function
Public Function simple_dbase_for_angle_(re As record_data_type) As Byte
Dim temp_record As total_record_type
Dim i%
For i% = 1 To last_subs_angle3_value
    temp_record.record_data = re
     Call add_conditions_to_record(angle3_value_, subs_angle3_value(i%).no, 0, 0, _
                   temp_record.record_data.data0.condition_data)
   Call remove_record(angle3_value_, subs_angle3_value(i%).no, 0)
   simple_dbase_for_angle_ = set_three_angle_value(subs_angle3_value(i%).A3_value.angle(0), _
           subs_angle3_value(i%).A3_value.angle(1), subs_angle3_value(i%).A3_value.angle(2), _
            subs_angle3_value(i%).A3_value.para(0), subs_angle3_value(i%).A3_value.para(1), _
             subs_angle3_value(i%).A3_value.para(2), _
              subs_angle3_value(i%).A3_value.value, False, temp_record, 0, 0, 0, 0, 0, 0, False)
   If simple_dbase_for_angle_ > 1 Then
      Exit Function
   End If
Next i%
End Function
Public Function change_poly_to_area_element(pol As polygon) As condition_type
If pol.total_v = 3 Then
 change_poly_to_area_element.ty = triangle_
 change_poly_to_area_element.no = triangle_number(pol.v(0), _
    pol.v(1), pol.v(2), 0, 0, 0, 0, 0, 0, 0)
ElseIf pol.total_v = 4 Then
 change_poly_to_area_element.ty = polygon_
 change_poly_to_area_element.no = polygon4_number(pol.v(0), _
   pol.v(1), pol.v(2), pol.v(3), 0)
End If
End Function

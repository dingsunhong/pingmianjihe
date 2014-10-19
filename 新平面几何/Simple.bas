Attribute VB_Name = "simple"
Public Sub simple_deline(ByVal i%)
Dim l%
Dim n(1) As Integer
l% = line_number0(Deline(i%).data(0).data0.poi(0), Deline(i%).data(0).data0.poi(1), _
         n(0), n(1))
If n(0) > n(1) Then
 l% = Deline(i%).data(0).data0.poi(0)
  Deline(i%).data(0).data0.poi(0) = Deline(i%).data(0).data0.poi(1)
     
End If
l% = line_number0(Deline(i%).data(0).data0.poi(2), Deline(i%).data(0).data0.poi(3), _
         n(0), n(1))
If n(0) > n(1) Then
 
  Deline(i%).data(0).data0.poi(2) = Deline(i%).data(0).data0.poi(3)
   
End If

End Sub
Public Function simple_four_point(ByVal n1%, ByVal n2%, ByVal n3%, _
      ByVal n4%, _
           on1%, on2%, on3%, on4%) As Integer
Dim tl(1) As Integer
If n1% = n3% Then
 If n2 < n4% Then
  on1% = n1%
   on2% = n2%
    on3% = n2%
     on4% = n4%
  simple_four_point = 1
 ElseIf n2 > n4% Then
 
  on1% = n3%
   on2% = n4%
    on3% = n4%
     on4% = n2%
   simple_four_point = 2
 End If
ElseIf n2% = n4% Then
 If n1% < n3% Then
  on1% = n1%
   on2% = n3%
    on3% = n3%
     on4% = n2%
    simple_four_point = 3
 ElseIf n1% > n3% Then
  on1% = n3%
   on2% = n1%
    on3% = n1%
     on4% = n2%
    simple_four_point = 4
 End If
ElseIf n2% = n3% Then
 on1% = n1%
  on2% = n2%
   on3% = n3%
    on4% = n4%
      simple_four_point = 5
 ElseIf n1% = n4% Then
  on1% = n3%
   on2% = n4%
    on3% = n4%
     on4% = n2%
     simple_four_point = 6
 Else
  If n1% < n3 Then
   on1% = n1%
    on2% = n2%
     on3% = n3%
      on4% = n4%
      simple_four_point = 7
  Else
   on1% = n3%
    on2% = n4%
     on3% = n1%
      on4% = n2%
       simple_four_point = 8
  End If
End If
End Function
Public Sub simple_point_pair(in_dp As point_pair_data0_type, _
                   out_dp As point_pair_data0_type)
Dim i%
Dim ty3 As Byte
Dim ty4 As Byte
Dim ty(1) As Boolean
Dim dp(1) As point_pair_data0_type
'Dim tp(7) As Integer
'Dim tn(7) As Integer
'Dim tl(3) As Integer
ty1 = 0
ty2 = 0
dp(0) = in_dp
out_dp = dp(1)
'************************************************************
'整理外项
'**************************************************************
Call simple_point_pair_item(dp(0).poi(0), dp(0).poi(1), dp(0).poi(2), dp(0).poi(3), _
       dp(0).poi(4), dp(0).poi(5), dp(0).poi(6), dp(0).poi(7), dp(0).n(0), _
        dp(0).n(1), dp(0).n(2), dp(0).n(3), dp(0).n(4), dp(0).n(5), dp(0).n(6), _
         dp(0).n(7), dp(0).line_no(0), dp(0).line_no(1), dp(0).line_no(2), dp(0).line_no(3))
'*********************************************************
'out_dp = dp(0)
dp(1) = dp(0)
'在两根线上
 Call simple_point_pair0(dp(0), dp(0), 0)
 'If dp(0).line_no(0) = dp(0).line_no(1) And dp(0).line_no(1) = dp(0).line_no(2) And dp(0).line_no(2) = dp(0).line_no(3) Then
  out_dp = dp(0)
 'End If
'If dp(0).line_no(0) <> dp(0).line_no(1) Or dp(0).line_no(2) <> dp(0).line_no(3) Then
 'out_dp = dp(1)
  'Exit Sub
'Else 'dp(0).line_no(0) = dp(0).line_no(1) And dp(0).line_no(2) = dp(0).line_no(3) Then
 ' If dp(0).line_no(1) <> dp(0).line_no(2) Then
  ' out_dp = dp(0)
   ' Exit Sub
  'Else
   'If dp(0).con_line_type(0) > 2 And dp(0).con_line_type(1) > 2 Then
    'out_dp = dp(0)
   'Else
    'call exchange_two_integer(dp(1).poi(2), dp(1).poi(4))
    'call exchange_two_integer(dp(1).poi(3), dp(1).poi(5))
    'call exchange_two_integer(dp(1).n(2), dp(1).n(4))
    'call exchange_two_integer(dp(1).n(3), dp(1).n(5))
    'Call simple_point_pair0(dp(1), dp(1), 0)
     'If dp(1).con_line_type(0) > 2 And dp(1).con_line_type(1) > 2 Then
      'out_dp = dp(1)
     'Else
      't_dp = dp(0)
     'End If
   'End If
  'End If
'End If
'simple_point_pair = True
End Sub

Public Function simple_two_triangle(p1%, p2%, p3%, p4%, p5%, p6%)
Dim t%
Dim tp1(5) As Integer
Dim tp2(5) As Integer
tp1(0) = p1%
tp1(1) = p2%
tp1(2) = p3%
tp1(3) = p4%
tp1(4) = p5%
tp1(5) = p6%
tp2(0) = p4%
tp2(1) = p5%
tp2(2) = p6%
tp2(3) = p1%
tp2(4) = p2%
tp2(5) = p3%
If tp1(0) > tp1(1) Then
 Call exchange_two_int(tp1(0), tp1(1))
  Call exchange_two_int(tp1(3), tp1(4))
End If
If tp1(1) > tp1(2) Then
 Call exchange_two_int(tp1(1), tp1(2))
  Call exchange_two_int(tp1(4), tp1(5))
End If
If tp1(0) > tp1(1) Then
 Call exchange_two_int(tp1(0), tp1(1))
  Call exchange_two_int(tp1(3), tp1(4))
End If
'**************************************
If tp2(0) > tp2(1) Then
 Call exchange_two_int(tp2(0), tp2(1))
  Call exchange_two_int(tp2(3), tp2(4))
End If
If tp2(1) > tp2(2) Then
 Call exchange_two_int(tp2(1), tp2(2))
  Call exchange_two_int(tp2(4), tp2(5))
End If
If tp2(0) > tp2(1) Then
 Call exchange_two_int(tp2(0), tp2(1))
  Call exchange_two_int(tp2(3), tp2(4))
End If
If tp1(0) < tp2(0) Then
 
ElseIf tp1(0) > tp2(0) Then
 
Else
 If tp1(1) < tp2(1) Then
  
 ElseIf tp1(1) > tp2(1) Then
  t% = 2
 Else
  If tp1(2) < tp2(2) Then
   
  ElseIf tp1(2) > tp2(2) Then
   t% = 2
  Else
  simple_two_triangle = False
  Exit Function
  End If
 End If
End If
  simple_two_triangle = True
If t% = 1 Then
 p1% = tp1(0)
 p2% = tp1(1)
 p3% = tp1(2)
 p4% = tp1(3)
 p5% = tp1(4)
 p6% = tp1(5)
Else
 p1% = tp2(0)
 p2% = tp2(1)
 p3% = tp2(2)
 p4% = tp2(3)
 p5% = tp2(4)
 p6% = tp2(5)
End If
End Function

Public Sub simple_point_pair_item(p1%, p2%, p3%, p4%, p5%, p6%, p7%, p8%, _
     n1%, n2%, n3%, n4%, n5%, n6%, n7%, n8%, l1%, l2%, l3%, l4%)
If l1% = l2% And l1% = l3% And l1% = l4% Then
 Exit Sub
End If
If (l1% > l4% Or (l1% = l4% And _
    (n1% > n7% Or (n1% = n7% And n2% > n8%)))) Or _
     (l4% = l2% And l1% <> l3%) Then
Call exchange_two_integer(l1%, l4%)
Call exchange_two_integer(p1%, p7%)
Call exchange_two_integer(p2%, p8%)
Call exchange_two_integer(n1%, n7%)
Call exchange_two_integer(n2%, n8%)
End If
If ((l2% > l3% Or (l2% = l3% And _
    (n3% > n5% Or (n3% = n5% And n4% > n6%)))) Or _
     (l1% = l3% And l2% <> l4%)) And (l1% <> l2%) Then
Call exchange_two_integer(l2%, l3%)
Call exchange_two_integer(p3%, p5%)
Call exchange_two_integer(p4%, p6%)
Call exchange_two_integer(n3%, n5%)
Call exchange_two_integer(n4%, n6%)
End If
If l1% <> l2% Then
If l2% = l4% Then
Call exchange_two_integer(l1%, l4%)
Call exchange_two_integer(p1%, p7%)
Call exchange_two_integer(p2%, p8%)
Call exchange_two_integer(n1%, n7%)
Call exchange_two_integer(n2%, n8%)
ElseIf l1% = l3% Then
Call exchange_two_integer(l2%, l3%)
Call exchange_two_integer(p3%, p5%)
Call exchange_two_integer(p4%, p6%)
Call exchange_two_integer(n3%, n5%)
Call exchange_two_integer(n4%, n6%)
ElseIf l3% = l4% Then
Call exchange_two_integer(l2%, l3%)
Call exchange_two_integer(p3%, p5%)
Call exchange_two_integer(p4%, p6%)
Call exchange_two_integer(n3%, n5%)
Call exchange_two_integer(n4%, n6%)
'*******
Call exchange_two_integer(l1%, l4%)
Call exchange_two_integer(p1%, p7%)
Call exchange_two_integer(p2%, p8%)
Call exchange_two_integer(n1%, n7%)
Call exchange_two_integer(n2%, n8%)
End If
End If
If (l1% > l2% Or (l1% = l2% And (n1% > n3% Or _
     (n1% = n3% And n2% > n4%)))) Or _
      (p1% = p3% And p2% = p4%) And (l3% > l4% Or _
       (l3% = l4% And (n5% > n7% Or _
        (n5% = n7% And n6% > n8%)))) Or _
          (l1% <> l2% And l3% = l4%) Then
Call exchange_two_integer(l1%, l2%)
Call exchange_two_integer(p1%, p3%)
Call exchange_two_integer(p2%, p4%)
Call exchange_two_integer(n1%, n3%)
Call exchange_two_integer(n2%, n4%)
Call exchange_two_integer(l3%, l4%)
Call exchange_two_integer(p5%, p7%)
Call exchange_two_integer(p6%, p8%)
Call exchange_two_integer(n5%, n7%)
Call exchange_two_integer(n6%, n8%)
End If
If l1% = l2% And l2% = l3% And l3% = l4% Then
 If n1% > n7 Then
 Call exchange_two_integer(l1%, l4%)
 Call exchange_two_integer(p1%, p7%)
 Call exchange_two_integer(p2%, p8%)
 Call exchange_two_integer(n1%, n7%)
 Call exchange_two_integer(n2%, n8%)
 End If
 If n3% > n5% Then
 Call exchange_two_integer(l2%, l3%)
 Call exchange_two_integer(p3%, p5%)
 Call exchange_two_integer(p4%, p6%)
 Call exchange_two_integer(n3%, n5%)
 Call exchange_two_integer(n4%, n6%)
 End If
 If n1% > n3% Then
 Call exchange_two_integer(l1%, l2%)
 Call exchange_two_integer(p1%, p3%)
 Call exchange_two_integer(p2%, p4%)
 Call exchange_two_integer(n1%, n3%)
 Call exchange_two_integer(n2%, n4%)
 Call exchange_two_integer(l3%, l4%)
 Call exchange_two_integer(p5%, p7%)
 Call exchange_two_integer(p6%, p8%)
 Call exchange_two_integer(n5%, n7%)
 Call exchange_two_integer(n6%, n8%)
 End If
End If
End Sub
Public Sub simple_point_pair0(in_dp As point_pair_data0_type, o_dp As point_pair_data0_type, _
              simple_times As Integer)
Dim dp(2) As point_pair_data0_type
Dim ty(2) As Boolean
dp(0) = in_dp
o_dp = dp(1) ' 初始化
dp(1) = dp(0)
dp(2) = dp(0)
Call exchange_two_integer(dp(2).poi(2), dp(2).poi(4))
Call exchange_two_integer(dp(2).poi(3), dp(2).poi(5))
Call exchange_two_integer(dp(2).n(2), dp(2).n(4))
Call exchange_two_integer(dp(2).n(3), dp(2).n(5))
Call exchange_two_integer(dp(2).line_no(1), dp(2).line_no(2))
ty(0) = arrange_four_point(dp(0).poi(0), dp(0).poi(1), dp(0).poi(2), dp(0).poi(3), _
                           dp(0).n(0), dp(0).n(1), dp(0).n(2), dp(0).n(3), dp(0).line_no(0), _
                             dp(0).line_no(1), dp(0).poi(0), dp(0).poi(1), dp(0).poi(2), _
                              dp(0).poi(3), 0, 0, dp(0).n(0), dp(0).n(1), dp(0).n(2), dp(0).n(3), _
                               0, 0, dp(0).line_no(0), dp(0).line_no(1), 0, dp(0).con_line_type(0), condition_data0, 0)
ty(1) = arrange_four_point(dp(0).poi(4), dp(0).poi(5), dp(0).poi(6), dp(0).poi(7), _
                           dp(0).n(4), dp(0).n(5), dp(0).n(6), dp(0).n(7), dp(0).line_no(2), _
                             dp(0).line_no(3), dp(0).poi(4), dp(0).poi(5), dp(0).poi(6), _
                              dp(0).poi(7), 0, 0, dp(0).n(4), dp(0).n(5), dp(0).n(6), dp(0).n(7), _
                               0, 0, dp(0).line_no(2), dp(0).line_no(3), 0, dp(0).con_line_type(1), condition_data0, 0)
If ty(0) = False Or ty(1) = False Then
ty(0) = arrange_four_point(dp(2).poi(0), dp(2).poi(1), dp(2).poi(2), dp(2).poi(3), _
                           dp(2).n(0), dp(2).n(1), dp(2).n(2), dp(2).n(3), dp(2).line_no(0), _
                             dp(2).line_no(1), dp(2).poi(0), dp(2).poi(1), dp(2).poi(2), _
                              dp(2).poi(3), 0, 0, dp(2).n(0), dp(2).n(1), dp(2).n(2), dp(2).n(3), _
                               0, 0, dp(2).line_no(0), dp(2).line_no(1), 0, dp(2).con_line_type(0), condition_data0, 0)
ty(1) = arrange_four_point(dp(2).poi(4), dp(2).poi(5), dp(2).poi(6), dp(2).poi(7), _
                           dp(2).n(4), dp(2).n(5), dp(2).n(6), dp(2).n(7), dp(2).line_no(2), _
                             dp(2).line_no(3), dp(2).poi(4), dp(2).poi(5), dp(2).poi(6), _
                              dp(2).poi(7), 0, 0, dp(2).n(4), dp(2).n(5), dp(2).n(6), dp(2).n(7), _
                               0, 0, dp(2).line_no(2), dp(2).line_no(3), 0, dp(2).con_line_type(1), condition_data0, 0)
ty(2) = True
If ty(0) = True And ty(1) = True Then
dp(0) = dp(2)
End If
End If
If dp(0).con_line_type(0) = dp(0).con_line_type(1) And dp(0).con_line_type(0) > 2 Then
    dp(0).con_line_type(0) = 3
    dp(0).con_line_type(1) = 3
    dp(0).poi(8) = dp(0).poi(0)
    dp(0).poi(9) = dp(0).poi(3)
    dp(0).poi(10) = dp(0).poi(4)
    dp(0).poi(11) = dp(0).poi(7)
    dp(0).n(8) = dp(0).n(0)
    dp(0).n(9) = dp(0).n(3)
    dp(0).n(10) = dp(0).n(4)
    dp(0).n(11) = dp(0).n(7)
    dp(0).line_no(4) = dp(0).line_no(0)
    dp(0).line_no(5) = dp(0).line_no(2)
    'o_dp = dp(0)
ElseIf (dp(0).con_line_type(0) = 3 And dp(0).con_line_type(1) = 5) Or _
             (dp(0).con_line_type(0) = 5 And dp(0).con_line_type(1) = 3) Or _
       (dp(0).con_line_type(0) = 4 And dp(0).con_line_type(1) = 8) Or _
             (dp(0).con_line_type(0) = 8 And dp(0).con_line_type(1) = 4) Or _
       (dp(0).con_line_type(0) = 6 And dp(0).con_line_type(1) = 7 Or _
             dp(0).con_line_type(0) = 7 And dp(0).con_line_type(1) = 6) Then
       Call exchange_two_integer(dp(0).poi(4), dp(0).poi(6))
       Call exchange_two_integer(dp(0).poi(5), dp(0).poi(7))
       Call exchange_two_integer(dp(0).n(4), dp(0).n(6))
       Call exchange_two_integer(dp(0).n(5), dp(0).n(7))
       dp(0).con_line_type(0) = 3
       dp(0).con_line_type(1) = 5
       dp(0).poi(8) = dp(0).poi(0)
       dp(0).poi(9) = dp(0).poi(3)
       dp(0).poi(10) = dp(0).poi(5)
       dp(0).poi(11) = dp(0).poi(6)
       dp(0).n(8) = dp(0).n(0)
       dp(0).n(9) = dp(0).n(3)
       dp(0).n(10) = dp(0).n(5)
       dp(0).n(11) = dp(0).n(6)
       dp(0).line_no(4) = dp(0).line_no(0)
       dp(0).line_no(5) = dp(0).line_no(2)
       'o_dp = dp(0)
Else
       dp(1).con_line_type(0) = dp(0).con_line_type(0)
       dp(1).con_line_type(1) = dp(0).con_line_type(1)
       dp(0) = dp(1)
End If
'ElseIf dp(0).con_line_type(1) = 5 And dp(0).con_line_type(0) > 2 And _
              dp(0).line_no(0) = dp(0).line_no(2) Then
 'If dp(0).con_line_type(0) < 9 Then
 'dp(0).poi(0) = dp(0).poi(4)
 'dp(0).poi(1) = dp(0).poi(5)
 'dp(0).poi(2) = dp(0).poi(6)
 'dp(0).poi(3) = dp(0).poi(7)
 'dp(0).n(0) = dp(0).n(4)
 'dp(0).n(1) = dp(0).n(5)
 'dp(0).n(2) = dp(0).n(6)
 'dp(0).n(3) = dp(0).n(7)
 'dp(0).line_no(0) = dp(0).line_no(2)
 'dp(0).line_no(1) = dp(0).line_no(3)
'**
 'dp(0).poi(4) = dp(1).poi(2)
 'dp(0).poi(5) = dp(1).poi(3)
 'dp(0).poi(6) = dp(1).poi(0)
 'dp(0).poi(7) = dp(1).poi(1)
 'dp(0).n(4) = dp(0).n(2)
 'dp(0).n(5) = dp(0).n(3)
 'dp(0).n(6) = dp(0).n(0)
 'dp(0).n(7) = dp(0).n(1)
 'dp(0).line_no(0) = dp(0).line_no(2)
 'dp(0).line_no(1) = dp(0).line_no(3)
' dp(0).con_line_type(0) = 0
 'dp(0).con_line_type(1) = 0
'Else
 'dp(0) = dp(1)
'End If
'Else
 'dp(0) = dp(1)
'End If
  If dp(0).line_no(0) = dp(0).line_no(1) And dp(0).line_no(0) = dp(0).line_no(2) And _
       dp(0).line_no(0) = dp(0).line_no(3) And simple_times = 0 And ty(2) = False Then
    dp(2) = dp(0)
     dp(0) = dp(1)
   Call exchange_two_integer(dp(0).poi(2), dp(0).poi(4))
   Call exchange_two_integer(dp(0).poi(3), dp(0).poi(5))
   Call exchange_two_integer(dp(0).n(2), dp(0).n(4))
   Call exchange_two_integer(dp(0).n(3), dp(0).n(5))
   Call simple_point_pair0(dp(0), dp(0), 1)
     If dp(2).con_line_type(0) > 2 And dp(0).con_line_type(0) > 2 Then
      Call simple_dpoint_pair_(dp(2))
      Call simple_dpoint_pair_(dp(0))
      If compare_two_point_pair(dp(2), dp(0), 0) >= 0 Then
       dp(0) = dp(2)
      End If
        'Call simple_dpoint_pair_(dp(0))
     ElseIf dp(2).con_line_type(0) > 2 Then
       dp(0) = dp(2)
        Call simple_dpoint_pair_(dp(0))
     Else
        Call simple_dpoint_pair_(dp(0))
     End If
        o_dp = dp(0)
  Else
       o_dp = dp(0)
  End If
End Sub

Public Function simple_two_two_triangle(t_triA1 As two_triangle_type, _
             T_triA2 As two_triangle_type, re1 As record_data_type, _
                 re2 As record_data_type, no_reduce As Byte) As Byte
                  '两次推出全等
Dim temp_record As record_data_type
Dim dir As Integer
Dim tri As triangle_data0_type
Dim i%
temp_record = re1
For i% = 1 To re2.data0.condition_data.condition_no
Call add_conditions_to_record(re2.data0.condition_data.condition(i%).ty, _
        re2.data0.condition_data.condition(i%).no, 0, 0, temp_record.data0.condition_data)
Next i%
     dir = set_direction(t_triA1.direction, T_triA2.direction)
      simple_two_two_triangle = _
        simple_two_two_triangle_(triangle(t_triA1.triangle(1)).data(0), _
           dir, temp_record, no_reduce)
      If simple_two_two_triangle > 1 Then
        Exit Function
      End If
     dir = set_direction(set_direction(t_triA1.direction, 1), _
               set_direction(T_triA2.direction, 1))
      simple_two_two_triangle = _
        simple_two_two_triangle_(triangle(t_triA1.triangle(0)).data(0), _
           dir, temp_record, no_reduce)
      If simple_two_two_triangle > 1 Then
        Exit Function
      End If
End Function

Public Function simple_two_two_triangle_(triA_ As triangle_data0_type, _
                      dir As Integer, re As record_data_type, _
                       no_reduce As Byte) As Byte
Dim temp_record As total_record_type
Dim pol As polygon
Dim triA As triangle_data0_type
triA = triA_
'两次全等的方向不同'成为等腰三角形
temp_record.record_data = re
If dir = -1 Then
simple_two_two_triangle_ = set_equal_dline(triA.poi(0), _
  triA.poi(1), triA.poi(0), triA.poi(2), 0, 0, 0, 0, 0, 0, _
   0, temp_record, 0, 0, 0, 0, no_reduce, False)
If simple_two_two_triangle_ > 1 Then
 Exit Function
End If
simple_two_two_triangle_ = set_three_angle_value(triA.angle(1), _
  triA.angle(2), 0, "1", "-1", "0", "0", 0, temp_record, 0, 0, 0, _
    no_reduce, 0, 0, False)
If simple_two_two_triangle_ > 1 Then
 Exit Function
End If
ElseIf dir = -2 Then
simple_two_two_triangle_ = set_equal_dline(triA.poi(2), _
  triA.poi(1), triA.poi(2), triA.poi(0), 0, 0, 0, 0, 0, 0, _
   0, temp_record, 0, 0, 0, 0, no_reduce, False)
If simple_two_two_triangle_ > 1 Then
 Exit Function
End If
simple_two_two_triangle_ = set_three_angle_value(triA.angle(0), _
  triA.angle(1), 0, "1", "-1", "0", "0", 0, temp_record, 0, 0, 0, _
     no_reduce, 0, 0, False)
If simple_two_two_triangle_ > 1 Then
 Exit Function
End If
ElseIf dir = -3 Then
simple_two_two_triangle_ = set_equal_dline(triA.poi(1), _
  triA.poi(0), triA.poi(1), triA.poi(2), 0, 0, 0, 0, 0, 0, _
   0, temp_record, 0, 0, 0, 0, no_reduce, False)
If simple_two_two_triangle_ > 1 Then
 Exit Function
End If
simple_two_two_triangle_ = set_three_angle_value(triA.angle(0), _
  triA.angle(2), 0, "1", "-1", "0", "0", 0, temp_record, 0, 0, 0, _
    no_reduce, 0, 0, False)
If simple_two_two_triangle_ > 1 Then
 Exit Function
End If
ElseIf dir = 2 Or dir = 3 Then
  pol.total_v = 3
  pol.v(0) = triA.poi(0)
  pol.v(1) = triA.poi(1)
  pol.v(2) = triA.poi(2)
  simple_two_two_triangle_ = set_Epolygon(pol, temp_record, 0, _
      no_reduce, 1, 0, False)
End If


End Function

Public Sub simple_dpoint_pair_(dp As point_pair_data0_type)
Dim ty As Byte
simple_point_pair_re:
      If (dp.n(0) > dp.n(2)) Or (dp.n(0) = dp.n(2) And dp.n(1) > dp.n(3)) Then
       Call exchange_two_integer(dp.poi(0), dp.poi(2))
       Call exchange_two_integer(dp.poi(1), dp.poi(3))
       Call exchange_two_integer(dp.poi(4), dp.poi(6))
       Call exchange_two_integer(dp.poi(5), dp.poi(7))
       Call exchange_two_integer(dp.n(0), dp.n(2))
       Call exchange_two_integer(dp.n(1), dp.n(3))
       Call exchange_two_integer(dp.n(4), dp.n(6))
       Call exchange_two_integer(dp.n(5), dp.n(7))
     GoTo simple_point_pair_re
      ElseIf dp.con_line_type(0) = 0 And _
       ((dp.n(0) > dp.n(6)) Or (dp.n(0) = dp.n(6) And dp.n(1) > dp.n(7))) Then
       Call exchange_two_integer(dp.poi(0), dp.poi(6))
       Call exchange_two_integer(dp.poi(1), dp.poi(7))
       Call exchange_two_integer(dp.n(0), dp.n(6))
       Call exchange_two_integer(dp.n(1), dp.n(7))
     GoTo simple_point_pair_re
      ElseIf (dp.n(0) > dp.n(4)) Or (dp.n(0) = dp.n(4) And dp.n(1) > dp.n(5)) Then
       Call exchange_two_integer(dp.poi(0), dp.poi(4))
       Call exchange_two_integer(dp.poi(1), dp.poi(5))
       Call exchange_two_integer(dp.poi(2), dp.poi(6))
       Call exchange_two_integer(dp.poi(3), dp.poi(7))
       Call exchange_two_integer(dp.n(0), dp.n(4))
       Call exchange_two_integer(dp.n(1), dp.n(5))
       Call exchange_two_integer(dp.n(2), dp.n(6))
       Call exchange_two_integer(dp.n(3), dp.n(7))
       ty = dp.con_line_type(0)
       dp.con_line_type(0) = dp.con_line_type(1)
       dp.con_line_type(1) = ty
       Call exchange_two_integer(dp.poi(8), dp.poi(10))
       Call exchange_two_integer(dp.poi(9), dp.poi(11))
       Call exchange_two_integer(dp.n(8), dp.n(10))
       Call exchange_two_integer(dp.n(9), dp.n(11))
     ElseIf dp.con_line_type(0) = 0 And _
       ((dp.n(2) > dp.n(4)) Or (dp.n(2) = dp.n(4) And dp.n(3) > dp.n(5))) Then
       Call exchange_two_integer(dp.poi(2), dp.poi(4))
       Call exchange_two_integer(dp.poi(3), dp.poi(5))
       Call exchange_two_integer(dp.n(2), dp.n(4))
       Call exchange_two_integer(dp.n(3), dp.n(5))
     GoTo simple_point_pair_re
      End If

End Sub
Public Function simple_item(ByVal n%) As Integer
Dim tn(1) As Integer
Dim tp(3) As Integer
Dim re_condition As condition_data_type
Dim temp_ele1() As element_data_type
Dim temp_ele2() As element_data_type
Dim last_temp_ele1%, last_temp_ele2%
If read_element_from_temp_item(n%, temp_ele1(), last_temp_ele1%, _
          temp_ele2(), last_temp_ele2%) = 1 Then
 While i% < last_temp_ele1%
  While j% < last_temp_ele2%
   If temp_ele1(i%).poi(0) = temp_ele2(j%).poi(0) And _
        temp_ele1(i%).poi(1) = temp_ele2(j%).poi(1) Then
         last_temp_ele1% = last_temp_ele1% - 1
          For k% = i% To last_temp_ele1%
           temp_ele1(k%) = temp_ele1(k% + 1)
          Next k%
         last_temp_ele2% = last_temp_ele2% - 1
          For k% = i% To last_temp_ele2%
           temp_ele2(k%) = temp_ele2(k% + 1)
          Next k%
         i% = i% - 1
    Else
     j% = j% + 1
    End If
  Wend
   i% = i% + 1
 Wend
 If last_temp_ele1% = 0 And last_temp_ele2% = 0 Then
  simple_item = 0
 ElseIf last_temp_ele1% <= 1 And last_temp_ele2% <= 1 Then
  If last_temp_ele1% = 1 And last_temp_ele2% = 1 Then
      Call set_item0(temp_ele1(0).poi(0), temp_ele1(0).poi(1), _
            temp_ele2(0).poi(0), temp_ele2(0).poi(1), _
             "/", 0, 0, 0, 0, 0, 0, item0(i%).data(0).para(0), item0(i%).data(0).para(1), _
              "1", "", "1", 0, re_condition, _
               0, simple_item, 0, 0, condition_data0, False)
  ElseIf last_temp_ele1% = 1 Then
      Call set_item0(temp_ele1(0).poi(0), temp_ele1(0).poi(1), _
            0, 0, "~", 0, 0, 0, 0, 0, 0, item0(i%).data(0).para(0), item0(i%).data(0).para(1), _
              "1", "", "1", 0, re_condition, _
               0, simple_item, 0, 0, condition_data0, False)
  ElseIf last_temp_ele2% = 1 Then
      Call set_item0(0, 0, temp_ele2(0).poi(0), temp_ele2(0).poi(1), _
             "/", 0, 0, 0, 0, 0, 0, item0(i%).data(0).para(0), item0(i%).data(0).para(1), _
              "1", "", "1", 0, re_condition, _
               0, simple_item, 0, 0, condition_data0, False)
  End If
 Else
  If last_temp_ele1% = 0 Then
   tp(0) = 0
   tp(1) = 0
  ElseIf last_temp_ele1% = 1 Then
   tp(0) = temp_ele1(0).poi(0)
   tp(1) = temp_ele1(0).poi(1)
  Else
   tp(0) = from_element_to_item0(temp_ele1(), last_temp_ele1%)
   If tp(0) > 0 Then
      tp(1) = -7
   Else
      tp(1) = 0
   End If
  End If
  If last_temp_ele2% = 0 Then
   tp(2) = 0
   tp(3) = 0
  ElseIf last_temp_ele2% = 1 Then
   tp(2) = temp_ele2(0).poi(0)
   tp(3) = temp_ele2(0).poi(1)
  Else
   tp(2) = from_element_to_item0(temp_ele2(), last_temp_ele2%)
   If tp(2) > 0 Then
      tp(3) = -7
   Else
      tp(3) = 0
   End If
  End If
  If tp(0) > 0 And tp(2) > 0 Then
        Call set_item0(tp(0), tp(1), tp(2), tp(3), item0(i%).data(0).sig, 0, 0, 0, 0, 0, 0, _
          item0(i%).data(0).para(0), item0(i%).data(0).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_item, 0, 0, condition_data0, False)
  ElseIf tp(0) > 0 Then
        Call set_item0(tp(0), tp(1), 0, 0, "~", 0, 0, 0, 0, 0, 0, _
          item0(i%).data(0).para(0), item0(i%).data(0).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_item, 0, 0, condition_data0, False)
  ElseIf tp(2) > 0 Then
   If item0(i%).data(0).sig = "*" Then
        Call set_item0(tp(2), tp(3), 0, 0, "~", 0, 0, 0, 0, 0, 0, _
          item0(i%).data(0).para(0), item0(i%).data(0).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_item, 0, 0, condition_data0, False)
   Else
        Call set_item0(0, 0, tp(2), tp(3), item0(i%).data(0).sig, 0, 0, 0, 0, 0, 0, _
          item0(i%).data(0).para(0), item0(i%).data(0).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_item, 0, 0, condition_data0, False)
   End If
  Else
   simple_item = 0
  End If
 End If
Else
If item0(n%).data(0).poi(1) = -7 And item0(n%).data(0).poi(3) = -7 Then
   If item0(item0(n%).data(0).poi(0)).data(0).sig = "~" And item0(item0(n%).data(0).poi(2)).data(0).sig = "~" Then
    Call set_item0(item0(item0(n%).data(0).poi(0)).data(0).poi(0), item0(item0(n%).data(0).poi(0)).data(0).poi(1), _
            item0(item0(n%).data(0).poi(2)).data(0).poi(0), item0(item0(n%).data(0).poi(2)).data(0).poi(1), _
             item0(n%).data(0).sig, 0, 0, 0, 0, 0, 0, item0(n%).data(0).para(0), item0(n%).data(0).para(1), _
              "1", "", "1", 0, re_condition, _
               0, simple_item, 0, 0, condition_data0, False)
   ElseIf item0(item0(n%).data(0).poi(0)).data(0).sig = "~" And item0(n%).data(0).sig <> "~" Then
    tn(0) = simple_item(item0(n%).data(0).poi(2))
     If item0(n%).data(0).sig = "~" Then
      Call set_item0(item0(item0(n%).data(0).poi(0)).data(0).poi(0), item0(item0(n%).data(0).poi(0)).data(0).poi(1), _
            item0(tn(0)).data(0).poi(0), item0(tn(0)).data(0).poi(1), _
             item0(n%).data(0).sig, 0, 0, 0, 0, 0, 0, item0(n%).data(0).para(0), item0(n%).data(0).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_item, 0, 0, condition_data0, False)
     Else
      Call set_item0(item0(item0(n%).data(0).poi(0)).data(0).poi(0), item0(item0(n%).data(0).poi(0)).data(0).poi(1), _
            tn(0), -7, item0(n%).data(0).sig, 0, 0, 0, 0, 0, 0, item0(n%).data(0).para(0), item0(n%).data(0).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_item, 0, 0, condition_data0, False)
     End If
   ElseIf item0(item0(n%).data(0).poi(2)).data(0).sig = "~" And item0(n%).data(0).sig <> "~" Then
    tn(0) = simple_item(item0(n%).data(0).poi(0))
     If item0(n%).data(0).sig = "~" Then
      Call set_item0(item0(tn(0)).data(0).poi(0), item0(tn(0)).data(0).poi(1), _
          item0(item0(n%).data(0).poi(2)).data(0).poi(0), item0(item0(n%).data(0).poi(2)).data(0).poi(1), _
             item0(n%).data(0).sig, 0, 0, 0, 0, 0, 0, item0(n%).data(0).para(0), item0(n%).data(0).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_item, 0, 0, condition_data0, False)
     Else
      Call set_item0(tn(0), -7, _
          item0(item0(n%).data(0).poi(2)).data(0).poi(0), item0(item0(n%).data(0).poi(2)).data(0).poi(1), _
           item0(n%).data(0).sig, 0, 0, 0, 0, 0, 0, item0(n%).data(0).para(0), item0(n%).data(0).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_item, 0, 0, condition_data0, False)
     End If
   End If
ElseIf item0(n%).data(0).poi(1) = -7 Then
 tn(0) = simple_item(item0(n%).data(0).poi(0))
  If item0(tn(0)).data(0).sig = "~" Then
  tp(0) = item0(tn(0)).data(0).poi(0)
  tp(1) = item0(tn(0)).data(0).poi(1)
  Else
  tp(0) = tn(0)
  tp(1) = -7
  End If
 If item0(item0(n%).data(0).poi(0)).data(0).sig = "~" Then
  
 Else
 End If
ElseIf item0(n%).data(0).poi(3) = -7 Then
 tn(0) = simple_item(item0(n%).data(0).poi(2))
  If item0(tn(0)).data(0).sig = "~" Then
  tp(0) = item0(tn(0)).data(0).poi(0)
  tp(1) = item0(tn(0)).data(0).poi(1)
  Else
  tp(0) = tn(0)
  tp(1) = -7
  End If
 If item0(item0(n%).data(0).poi(2)).data(0).sig = "~" Then
      Call set_item0(item0(tn(0)).data(0).poi(0), item0(tn(0)).data(0).poi(1), _
          item0(item0(n%).data(0).poi(2)).data(0).poi(0), item0(item0(n%).data(0).poi(2)).data(0).poi(1), _
         item0(n%).data(0).sig, 0, 0, 0, 0, 0, 0, item0(n%).data(0).para(0), item0(n%).data(0).para(1), _
              "1", "", "1", 0, re_condition, 0, simple_item, 0, 0, condition_data0, False)
 Else
 End If
Else
 tn(0) = simple_item(item0(n%).data(0).poi(0))
   If item0(tn(0)).data(0).sig = "~" Then
    tp(0) = item0(tn(0)).data(0).poi(0)
    tp(1) = item0(tn(0)).data(0).poi(1)
   Else
    tp(0) = tn(0)
    tp(1) = -7
   End If
  tn(1) = simple_item(item0(n%).data(0).poi(1))
   If item0(tn(0)).data(0).sig = "~" Then
    tp(2) = item0(tn(1)).data(0).poi(0)
    tp(3) = item0(tn(1)).data(0).poi(1)
   Else
    tp(0) = tn(1)
    tp(1) = -7
   End If
 End If
End If
End Function
Public Function read_element_from_item(ByVal n%, ele1() As element_data_type, last_ele1%, _
              ele2() As element_data_type, last_ele2%) As Byte
Dim temp_ele1() As element_data_type
Dim last_temp_ele1%
Dim temp_ele2() As element_data_type
Dim last_temp_ele2%
Dim i%
If item0(n%).data(0).sig = "~" Then
ReDim Preserve ele1(last_ele1%) As element_data_type
ele1(last_ele1%).poi(0) = item0(n%).data(0).poi(0)
ele1(last_ele1%).poi(1) = item0(n%).data(0).poi(1)
last_ele1% = last_ele1% + 1
read_element_from_item = 1
ElseIf item0(n%).data(0).sig = "*" Then
 If item0(n%).data(0).poi(1) <> -7 And item0(n%).data(0).poi(3) <> -7 Then
  ReDim Preserve ele1(last_ele1%) As element_data_type
  ele1(last_ele1%).poi(0) = item0(n%).data(0).poi(0)
  ele1(last_ele1%).poi(1) = item0(n%).data(0).poi(1)
  last_ele1% = last_ele1% + 1
  ReDim Preserve ele1(last_ele1%) As element_data_type
  ele1(last_ele1%).poi(0) = item0(n%).data(0).poi(2)
  ele1(last_ele1%).poi(1) = item0(n%).data(0).poi(3)
  last_ele1% = last_ele1% + 1
  read_element_from_item = 1
 ElseIf item0(n%).data(0).poi(1) = -7 Then
  If read_element_from_item(item0(n%).data(0).poi(0), temp_ele1(), last_temp_ele1%, _
        temp_ele2(), last_temp_ele2%) = 1 Then
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%).poi(0) = item0(n%).data(0).poi(2)
        ele1(last_ele1%).poi(1) = item0(n%).data(0).poi(3)
       last_ele1% = last_ele1% + 1
     For i% = 1 To last_temp_ele1% - 1
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%) = temp_ele1(i%)
        last_ele1% = last_ele1% + 1
     Next i%
     For i% = 1 To last_temp_ele2% - 1
      ReDim Preserve ele2(last_ele2%) As element_data_type
       ele2(last_ele2%) = temp_ele2(i%)
        last_ele2% = last_ele2% + 1
     Next i%
      read_element_from_item = 1
  Else
   read_element_from_item = 0
  End If
 ElseIf item0(n%).data(0).poi(3) = -7 Then
  If read_element_from_item(item0(n%).data(0).poi(2), temp_ele1(), last_temp_ele1%, _
        temp_ele2(), last_temp_ele2%) = 1 Then
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%).poi(0) = item0(n%).data(0).poi(0)
        ele1(last_ele1%).poi(1) = item0(n%).data(0).poi(1)
       last_ele1% = last_ele1% + 1
     For i% = 1 To last_temp_ele1% - 1
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%) = temp_ele1(i%)
        last_ele1% = last_ele1% + 1
     Next i%
     For i% = 1 To last_temp_ele2% - 1
      ReDim Preserve ele2(last_ele2%) As element_data_type
       ele2(last_ele2%) = temp_ele2(i%)
        last_ele2% = last_ele2% + 1
     Next i%
      read_element_from_item = 1
  Else
   read_element_from_item = 0
  End If
 Else
  If read_element_from_item(item0(n%).data(0).poi(0), temp_ele1(), last_temp_ele1%, _
        temp_ele2(), last_temp_ele2%) = 1 Then
    If read_element_from_item(item0(n%).data(0).poi(2), temp_ele1(), last_temp_ele1%, _
        temp_ele2(), last_temp_ele2%) = 1 Then
     For i% = 1 To last_temp_ele1% - 1
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%) = temp_ele1(i%)
        last_ele1% = last_ele1% + 1
     Next i%
     For i% = 1 To last_temp_ele2% - 1
      ReDim Preserve ele2(last_ele2%) As element_data_type
       ele1(last_ele2%) = temp_ele2(i%)
        last_ele2% = last_ele2% + 1
     Next i%
      read_element_from_item = 1
  Else
   read_element_from_item = 0
  End If
 Else
   read_element_from_item = 0
  End If
 End If
ElseIf item0(n%).data(0).sig = "/" Then
 If item0(n%).data(0).poi(1) <> -7 And item0(n%).data(0).poi(3) <> -7 Then
  ReDim Preserve ele1(last_ele1%) As element_data_type
  ele1(last_ele1%).poi(0) = item0(n%).data(0).poi(0)
  ele1(last_ele1%).poi(1) = item0(n%).data(0).poi(1)
  last_ele1% = last_ele1% + 1
  ReDim Preserve ele2(last_ele2%) As element_data_type
  ele2(last_ele1%).poi(0) = item0(n%).data(0).poi(2)
  ele2(last_ele1%).poi(1) = item0(n%).data(0).poi(3)
  last_ele2% = last_ele2% + 1
  read_element_from_item = 1
 ElseIf item0(n%).data(0).poi(1) = -7 Then
   If read_element_from_item(item0(n%).data(0).poi(0), temp_ele1(), last_temp_ele1%, _
        temp_ele2(), last_temp_ele2%) = 1 Then
      ReDim Preserve ele2(last_ele2%) As element_data_type
       ele2(last_ele2%).poi(0) = item0(n%).data(0).poi(2)
        ele2(last_ele2%).poi(1) = item0(n%).data(0).poi(3)
       last_ele2% = last_ele2% + 1
     For i% = 1 To last_temp_ele1% - 1
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%) = temp_ele1(i%)
        last_ele1% = last_ele1% + 1
     Next i%
     For i% = 1 To last_temp_ele2% - 1
      ReDim Preserve ele1(last_ele2%) As element_data_type
       ele2(last_ele2%) = temp_ele2(i%)
        last_ele2% = last_ele2% + 1
     Next i%
      read_element_from_item = 1
  Else
   read_element_from_item = 0
  End If
 ElseIf item0(n%).data(0).poi(3) = -7 Then
  If read_element_from_item(item0(n%).data(0).poi(2), temp_ele1(), last_temp_ele1%, _
        temp_ele2(), last_temp_ele2%) = 1 Then
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%).poi(0) = item0(n%).data(0).poi(0)
        ele1(last_ele1%).poi(1) = item0(n%).data(0).poi(1)
       last_ele1% = last_ele1% + 1
     For i% = 1 To last_temp_ele1% - 1
      ReDim Preserve ele2(last_ele2%) As element_data_type
       ele2(last_ele2%) = temp_ele1(i%)
        last_ele2% = last_ele2% + 1
     Next i%
     For i% = 1 To last_temp_ele2% - 1
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%) = temp_ele1(i%)
        last_ele1% = last_ele1% + 1
     Next i%
      read_element_from_item = 1
  Else
   read_element_from_item = 0
  End If
 Else
   If read_element_from_item(item0(n%).data(0).poi(0), temp_ele1(), last_temp_ele1%, _
        temp_ele2(), last_temp_ele2%) = 1 Then
    If read_element_from_item(item0(n%).data(0).poi(2), temp_ele2(), last_temp_ele2%, _
        temp_ele1(), last_temp_ele1%) = 1 Then
     For i% = 1 To last_temp_ele1% - 1
      ReDim Preserve ele1(last_ele1%) As element_data_type
       ele1(last_ele1%) = temp_ele1(i%)
        last_ele1% = last_ele1% + 1
     Next i%
     For i% = 1 To last_temp_ele2% - 1
      ReDim Preserve ele2(last_ele2%) As element_data_type
       ele2(last_ele2%) = temp_ele2(i%)
        last_ele2% = last_ele2% + 1
     Next i%
      read_element_from_item = 1
  Else
   read_element_from_item = 0
  End If
 Else
   read_element_from_item = 0
  End If
 End If
Else
  read_element_from_item = 0
End If
End Function

Public Function simple_equation(ByVal s As String, re As total_record_type, out_s As String) As Byte
Dim n0%, n1%, n2%, n4%, n3%, i%, k%, po%, no%, tn%, sig%
Dim ts(3) As String
Dim ts_(3) As String
Dim ch_$
Dim ch$
Dim ch1$
Dim ty As Byte
Dim ty_ As Byte
Dim is_x_in%
If do_factor1(s, ts(0), ts(1), ts(2), ts(3), tn%) = False Then '分解
 simple_equation = 0
  Exit Function
Else '分解因子
 If tn% > 1 Then
 '每个因子=0
 If is_contain_x(ts(0), "x", 1) Then
   simple_equation = simple_equation(ts(0), re, out_s)
    If simple_equation > 1 Then
        Exit Function
    End If
 End If
 If is_contain_x(ts(1), "x", 1) Then
     simple_equation = set_equation(ts(1), 0, re)
      If simple_equation > 1 Then
       Exit Function
      End If
 End If
 If is_contain_x(ts(2), "x", 1) Then
     simple_equation = set_equation(ts(2), 0, re)
      If simple_equation > 1 Then
       Exit Function
      End If
 End If
 If is_contain_x(ts(3), "x", 1) Then
    simple_equation = set_equation(ts(3), 0, re)
     If simple_equation > 1 Then
       Exit Function
      End If
 End If
Else '不 能分解
 Call remove_brace(s$)
 n0% = 1
 n1% = InStr(n0%, s, "[", 0) '第一个根号
simple_equation_back1:
 If n1% > 0 Then
 ch$ = read_sqr_no_from_string(s$, n1%, n2%, "") '第一个根号
  If InStr(1, ch$, "x", 0) = 0 Then ''第一个根号不含未知数次
     ch$ = ""
      n0% = n2% + 1
       n1% = 0
        n2% = 0
        GoTo simple_equation_back1 '重新取第一个根号
  End If
simple_equation_back2:
  n3% = InStr(n2% + 1, s$, "[", 0)  '第二个根号
   If n3% > 0 Then
    ch1$ = read_sqr_no_from_string(s$, n3%, n4%, "")
      If InStr(1, ch1$, "x", 0) = 0 Then '第一个根号不含未知数次
       ch1$ = ""
      n2% = n4%
        n3% = 0
         n4% = 0
        GoTo simple_equation_back2
      Else
        If InStr(n4% + 1, s$, "[", 0) > 0 Then '有三个根号
         out_s = ""
          simple_equation = 1
          no% = 0
           Exit Function
        End If
     End If
   Else
    ch1$ = ""
   End If
 If is_contain_x(ch$, "x", 1) = 0 And is_contain_x(ch1$, "x", 1) = 0 Then
   simple_equation = 0
    out_s = s
    Exit Function
 Else
 If n1% > 0 And n2% > 0 Then
 po = 1
 sig = 0
 For k% = 2 To Len(s)
  ch = Mid$(s, k%, 1)
  If ch = "+" Or ch = "-" Or ch = "@" Or ch = "#" Then
  sig = sig + 1 '读出符号
  no% = k%
   If po% <= n1% And n2% < no% Then '根号位于两符号间
    ch1$ = Mid$(s$, po%, no% - po%) '读出根号
      If ch1$ = "" Or (sig = 1 And InStr(no% + 2, s$, "+") = 0 And _
          InStr(no% + 2, s$, "-") = 0 And InStr(no% + 2, s, "#") = 0 And _
            InStr(no% + 2, s, "@") = 0) Then
     If po% > 1 Then
     ch_ = Mid$(s$, 1, po% - 1)
     Else
     ch_ = ""
     End If
     ch_ = ch_ + Mid$(s$, no%, Len(s$) - no% + 1)
    simple_equation = 0
    s$ = time_string(ch1$, ch1$, False, False)
    out_s = minus_string(s$, time_string(ch_, ch_, False, False), True, False)
    Else
     simple_equation = 1
      Exit Function
    End If
   Else
      po% = no%
   End If
   End If
  Next k%
  Else
   out_s = s$
    simple_equation = 0
  End If
 End If
Else '无根号
  simple_equation = 0
   out_s = s
   Exit Function
End If
End If
End If
End Function




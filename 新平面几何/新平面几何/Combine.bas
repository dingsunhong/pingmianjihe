Attribute VB_Name = "combine"
Option Explicit
Public Function combine_dpoint_pair_with_dpoint_pair(ByVal dp%, _
    ByVal no_reduce As Byte) As Byte
'合并两个比例线段式
Dim i%, j%, k%, l%, t_n%, t_m%, no%, ind%
Dim last_tn%, last_tn1%, last_tn2%
Dim n(5) As Integer
Dim m(5) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim tn1() As Integer
Dim tn2() As Integer
Dim ddp As point_pair_data0_type
'If Ddpoint_pair(dp%).record_.no_reduce > 4 Then
 '已简化
' Exit Function
'End If
For i% = 0 To 5 '=2=3 重复
'搜索相同线段
 If i% < 4 Then
  n(0) = i%
   n(1) = (i% + 1) Mod 4
    n(2) = (i% + 2) Mod 4
     n(3) = (i% + 3) Mod 4
 Else
  If Ddpoint_pair(dp%).data(0).data0.con_line_type(0) = 3 _
      And (Ddpoint_pair(dp%).data(0).data0.con_line_type(1) = 3 Or _
       Ddpoint_pair(dp%).data(0).data0.con_line_type(1) = 5) Then
   '前两项和后两项分段形式相同
   If i% = 4 Then
    n(0) = 4
     n(1) = 5
   Else
    n(0) = 5
     n(1) = 4
   End If
  Else
   GoTo combine_dpoint_pair_with_dpoint_pair_mark0
  End If
 End If
If Ddpoint_pair(dp%).data(0).data0.poi(2 * n(0)) > 0 And _
     Ddpoint_pair(dp%).data(0).data0.poi(2 * n(0) + 1) > 0 Then
 For j% = 0 To 5
 If j% < 4 Then
  m(0) = j%
   m(1) = (j% + 1) Mod 4
    m(2) = (j% + 2) Mod 4
     m(3) = (j% + 3) Mod 4
 Else
  If j% = 4 Then
   m(0) = 4
    m(1) = 5
  Else
   m(0) = 5
    m(1) = 4
  End If
 End If
 ddp.poi(2 * m(0)) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n(0))
 ddp.poi(2 * m(0) + 1) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n(0) + 1)
 ind% = j%
 ddp.poi(2 * m(1)) = -1
 ddp.poi(2 * m(1) + 1) = -2
Call search_for_point_pair(ddp, ind%, n_(0), 1)
 ddp.poi(2 * m(1)) = 30000
Call search_for_point_pair(ddp, ind%, n_(1), 1) '5.7
If n(0) = 1 Then
n(1) = 0
n(2) = 3
n(3) = 2
ElseIf n(0) = 3 Then
n(1) = 2
n(2) = 1
n(3) = 0
End If
If m(0) = 1 Then
m(1) = 0
m(2) = 3
m(3) = 2
ElseIf m(0) = 3 Then
m(1) = 2
m(2) = 1
m(3) = 0
End If
last_tn% = 0
last_tn1% = 0
last_tn2% = 0
For k% = n_(0) + 1 To n_(1)
no% = Ddpoint_pair(k%).data(0).record.data1.index.i(ind%)
If no% > 0 And no% < dp% And _
 Ddpoint_pair(no%).record_.no_reduce < 4 Then
  'If is_two_record_related(dpoint_pair_, no%, Ddpoint_pair(no%).data(0).record, _
     dpoint_pair_, dp%, Ddpoint_pair(dp%).data(0).record) = False Then
 If n(0) < 4 And m(0) < 4 Then
  last_tn% = last_tn% + 1
   ReDim Preserve tn(last_tn%) As Integer
    tn(last_tn%) = no%
 ElseIf n(0) > 3 And m(0) < 4 Then
  If Ddpoint_pair(dp%).data(0).data0.con_line_type(0) = 3 And _
      (Ddpoint_pair(dp%).data(0).data0.con_line_type(1) = 3 Or _
        Ddpoint_pair(dp%).data(0).data0.con_line_type(1) = 5) Then
   last_tn1% = last_tn1% + 1
   ReDim Preserve tn1(last_tn1%) As Integer
    tn1(last_tn1%) = no%
   End If
 ElseIf n(0) < 4 And m(0) > 3 Then
  If Ddpoint_pair(no%).data(0).data0.con_line_type(0) = 3 Or _
      (Ddpoint_pair(no%).data(0).data0.con_line_type(1) = 3 Or _
        Ddpoint_pair(dp%).data(0).data0.con_line_type(1) = 5) Then
   last_tn2% = last_tn2% + 1
    ReDim Preserve tn2(last_tn2%) As Integer
     tn2(last_tn2%) = no%
  End If
 End If
'End If
End If
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
     combine_dpoint_pair_with_dpoint_pair = _
     combine_dpoint_pair_with_dpoint_pair_(dp%, n(0), n(1), _
      n(2), n(3), no%, m(0), m(1), m(2), m(3), 0)
 If combine_dpoint_pair_with_dpoint_pair > 1 Then
  Exit Function
 End If
Next k%
For k% = 1 To last_tn1%
   no% = tn1(k%)
 If n(0) = 4 Then
    combine_dpoint_pair_with_dpoint_pair = _
     combine_dpoint_pair_with_dpoint_pair_(dp%, 4, 0, 5, 2, _
       no%, m(0), m(1), m(2), m(3), 0)
  If combine_dpoint_pair_with_dpoint_pair > 1 Then
   Exit Function
  End If
    combine_dpoint_pair_with_dpoint_pair = _
     combine_dpoint_pair_with_dpoint_pair_(dp%, 4, 1, 5, 3, _
       no%, m(0), m(1), m(2), m(3), 0)
 If combine_dpoint_pair_with_dpoint_pair > 1 Then
  Exit Function
 End If
 ElseIf n(0) = 5 Then
    combine_dpoint_pair_with_dpoint_pair = _
     combine_dpoint_pair_with_dpoint_pair_(dp%, 5, 2, 4, 0, _
       no%, m(0), m(1), m(2), m(3), 0)
 If combine_dpoint_pair_with_dpoint_pair > 1 Then
  Exit Function
 End If
    combine_dpoint_pair_with_dpoint_pair = _
     combine_dpoint_pair_with_dpoint_pair_(dp%, 5, 3, 4, 1, _
       no%, m(0), m(1), m(2), m(3), 0)
 If combine_dpoint_pair_with_dpoint_pair > 1 Then
  Exit Function
 End If
 End If
Next k%
For k% = 1 To last_tn2%
   no% = tn2(k%)
If m(0) = 4 Then
    combine_dpoint_pair_with_dpoint_pair = _
     combine_dpoint_pair_with_dpoint_pair_(dp%, n(0), n(1), _
      n(2), n(3), no%, 4, 0, 5, 2, 0)
 If combine_dpoint_pair_with_dpoint_pair > 1 Then
  Exit Function
 End If
    combine_dpoint_pair_with_dpoint_pair = _
     combine_dpoint_pair_with_dpoint_pair_(dp%, n(0), n(1), _
      n(2), n(3), no%, 4, 1, 5, 3, 0)
 If combine_dpoint_pair_with_dpoint_pair > 1 Then
  Exit Function
 End If
ElseIf m(0) = 5 Then
    combine_dpoint_pair_with_dpoint_pair = _
     combine_dpoint_pair_with_dpoint_pair_(dp%, n(0), n(1), _
      n(2), n(3), no%, 5, 2, 4, 0, 0)
 If combine_dpoint_pair_with_dpoint_pair > 1 Then
  Exit Function
 End If
    combine_dpoint_pair_with_dpoint_pair = _
     combine_dpoint_pair_with_dpoint_pair_(dp%, n(0), n(1), _
      n(2), n(3), no%, 5, 3, 4, 1, 0)
 If combine_dpoint_pair_with_dpoint_pair > 1 Then
  Exit Function
 End If
End If
Next k%
Next j%
End If
combine_dpoint_pair_with_dpoint_pair_mark0:
Next i%
End Function
Public Function combine_eline_with_line_value0(ByVal e%, ByVal l_v%, ByVal k%) As Byte
Dim temp_record As total_record_type
     temp_record.record_data.data0.condition_data.condition_no = 2
     temp_record.record_data.data0.condition_data.condition(1).ty = eline_
     temp_record.record_data.data0.condition_data.condition(1).no = e%
     temp_record.record_data.data0.condition_data.condition(2).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(2).no = l_v%
     temp_record.record_data.data0.theorem_no = 1
     If line_value(l_v%).data(0).data0.poi(0) = Deline(e%).data(0).data0.poi(2 * k%) Then '左端点重合
      If line_value(l_v%).data(0).data0.n(1) > Deline(e%).data(0).data0.n(2 * k% + 1) Then '等线短
       combine_eline_with_line_value0 = set_two_line_value(Deline(e%).data(0).data0.poi(2 * ((k% + 1) Mod 2)), _
         Deline(e%).data(0).data0.poi(2 * ((k% + 1) Mod 2) + 1), Deline(e%).data(0).data0.poi(2 * k% + 1), _
          line_value(l_v%).data(0).data0.poi(1), Deline(e%).data(0).data0.n(2 * ((k% + 1) Mod 2)), _
         Deline(e%).data(0).data0.n(2 * ((k% + 1) Mod 2) + 1), 0, 0, Deline(e%).data(0).data0.line_no((k% + 1) Mod 2), _
           0, "1", "1", line_value(l_v%).data(0).data0.value_, _
            temp_record, 0, 0)
      If combine_eline_with_line_value0 > 1 Then
       Exit Function
      End If
      ElseIf line_value(l_v%).data(0).data0.n(1) > Deline(e%).data(0).data0.n(2 * k% + 1) Then
      If line_value(l_v%).data(0).data0.n(1) < Deline(e%).data(0).data0.n(2 * k% + 1) Then
       combine_eline_with_line_value0 = set_two_line_value(Deline(e%).data(0).data0.poi(2 * ((k% + 1) Mod 2)), _
         Deline(e%).data(0).data0.poi(2 * ((k% + 1) Mod 2) + 1), Deline(e%).data(0).data0.poi(2 * k% + 1), _
          line_value(l_v%).data(0).data0.poi(1), Deline(e%).data(0).data0.n(2 * ((k% + 1) Mod 2)), _
         Deline(e%).data(0).data0.n(2 * ((k% + 1) Mod 2) + 1), 0, 0, Deline(e%).data(0).data0.line_no((k% + 1) Mod 2), _
           0, "-1", "1", line_value(l_v%).data(0).data0.value_, _
            temp_record, 0, 0)
      If combine_eline_with_line_value0 > 1 Then
       Exit Function
      End If
      End If
      End If
     ElseIf line_value(l_v%).data(0).data0.poi(1) = Deline(e%).data(0).data0.poi(2 * k% + 1) Then
      If line_value(l_v%).data(0).data0.n(0) < Deline(e%).data(0).data0.n(2 * k%) Then
      combine_eline_with_line_value0 = set_two_line_value(Deline(e%).data(0).data0.poi(2 * ((k% + 1) Mod 2)), _
         Deline(e%).data(0).data0.poi(2 * ((k% + 1) Mod 2) + 1), line_value(l_v%).data(0).data0.poi(0), _
          Deline(e%).data(0).data0.poi(2 * k%), Deline(e%).data(0).data0.n(2 * ((k% + 1) Mod 2)), _
         Deline(e%).data(0).data0.n(2 * ((k% + 1) Mod 2) + 1), 0, _
          0, Deline(e%).data(0).data0.line_no((k% + 1) Mod 2), _
           0, "1", "1", line_value(l_v%).data(0).data0.value_, _
            temp_record, 0, 0)
      If combine_eline_with_line_value0 > 1 Then
       Exit Function
      End If
     ElseIf line_value(l_v%).data(0).data0.n(0) > Deline(e%).data(0).data0.n(2 * k%) Then
      combine_eline_with_line_value0 = set_two_line_value(Deline(e%).data(0).data0.poi(2 * ((k% + 1) Mod 2)), _
         Deline(e%).data(0).data0.poi(2 * ((k% + 1) Mod 2) + 1), line_value(l_v%).data(0).data0.poi(0), _
          Deline(e%).data(0).data0.poi(2 * k%), Deline(e%).data(0).data0.n(2 * ((k% + 1) Mod 2)), _
         Deline(e%).data(0).data0.n(2 * ((k% + 1) Mod 2) + 1), 0, _
          0, Deline(e%).data(0).data0.line_no((k% + 1) Mod 2), _
           0, "1", "-1", line_value(l_v%).data(0).data0.value_, _
            temp_record, 0, 0)
      If combine_eline_with_line_value0 > 1 Then
       Exit Function
      End If
     
     End If
     End If
End Function

Public Function combine_relation_with_dpoint_pair0_(p() As Integer, _
          n_() As Integer, l_() As Integer, v() As String, _
           dp%, n1%, n2%, n3%, n4%, re As record_data_type, no_reduce As Byte) As Byte
Dim tn%
Dim it(1) As Integer
Dim para(1) As String
Dim temp_record As total_record_type
If v(0) <> "" Then
 If Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%) = p(2) And _
    Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1) = p(3) Then ' 有比项同
   temp_record.record_data = re
      combine_relation_with_dpoint_pair0_ = _
       combine_relation_with_dpoint_pair00_(dp%, _
        n1%, n2%, n3%, n4%, v(0), temp_record.record_data)
   If combine_relation_with_dpoint_pair0_ > 1 Then
    Exit Function
   End If
      combine_relation_with_dpoint_pair0_ = _
       combine_relation_with_h_point_pair0(v(0), Ddpoint_pair(dp%).data(0).data0, _
        n1%, n2%, n3%, n4%, temp_record)
   If combine_relation_with_dpoint_pair0_ > 1 Then
    Exit Function
   End If
 ElseIf Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%) = p(2) And _
    Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1) = p(3) Then
   temp_record.record_data = re
      combine_relation_with_dpoint_pair0_ = _
       combine_relation_with_dpoint_pair00_(dp%, _
        n1%, n3%, n2%, n4%, v(0), temp_record.record_data)
   If combine_relation_with_dpoint_pair0_ > 1 Then
    Exit Function
   End If
       combine_relation_with_dpoint_pair0_ = _
       combine_relation_with_h_point_pair0(v(0), Ddpoint_pair(dp%).data(0).data0, _
        n1%, n3%, n2%, n4%, temp_record)
   If combine_relation_with_dpoint_pair0_ > 1 Then
    Exit Function
   End If
ElseIf Ddpoint_pair(dp%).data(0).data0.poi(2 * n4%) = p(2) And _
    Ddpoint_pair(dp%).data(0).data0.poi(2 * n4% + 1) = p(3) Then
     temp_record.record_data = re
     If v(0) = "1" Then
     combine_relation_with_dpoint_pair0_ = set_dpoint_pair( _
      p(2), p(3), Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%), _
      Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1), Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%), _
      Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1), p(2), p(3), n_(2), n_(3), _
      Ddpoint_pair(dp%).data(0).data0.n(2 * n2%), Ddpoint_pair(dp%).data(0).data0.n(2 * n2% + 1), _
      Ddpoint_pair(dp%).data(0).data0.n(2 * n3%), Ddpoint_pair(dp%).data(0).data0.n(2 * n3% + 1), _
      n_(2), n_(3), l_(1), Ddpoint_pair(dp%).data(0).data0.line_no(n2%), _
      Ddpoint_pair(dp%).data(0).data0.line_no(n3%), l_(1), 1, _
       temp_record, False, 0, 0, 0, 0, False)
      If combine_relation_with_dpoint_pair0_ > 1 Then
       Exit Function
      End If
     Else '
      If Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%) And _
           Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1) Then
            combine_relation_with_dpoint_pair0_ = set_Drelation( _
             Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%), Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1), _
              p(2), p(3), 0, 0, 0, 0, 0, 0, sqr_string(v(0), True, False), temp_record, 0, 0, 0, 0, 0, False)
      If combine_relation_with_dpoint_pair0_ > 1 Then
       Exit Function
      End If
      Else
      'Call set_item0(p(2), p(3), p(2), p(3), "*", 0, 0, 0, 0, 0, 0, "1", "1", "1,", "", para(0), 0, _
               condition_data0, 0, it(0),   0)
      'Call set_item0(Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%), Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1), _
             Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%), Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1), "*", _
                0, 0, 0, 0, 0, 0, "1", "1", "1,", "", "", 0, _
               condition_data0, 0, it(0), 0)
       ' combine_relation_with_dpoint_pair0_ = set_general_string(it(0), it(1), 0, 0, v(0), "-1", "0", "0", "0", 0, _
           0, temp_record, 0, 0)
      'If combine_relation_with_dpoint_pair0_ > 1 Then
       'Exit Function
      'End If
     End If
     End If
 Else
    If v(0) = "1" Then
       temp_record.record_data = re
       If is_same_two_point(p(0), p(1), Ddpoint_pair(dp%).data(0).data0.poi(2 * n4%), _
         Ddpoint_pair(dp%).data(0).data0.poi(2 * n4% + 1)) Then
       combine_relation_with_dpoint_pair0_ = set_dpoint_pair( _
         p(2), p(3), Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%), _
           Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1), _
            Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%), _
             Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1), _
              p(2), p(3), n_(2), n_(3), Ddpoint_pair(dp%).data(0).data0.n(2 * n2%), _
           Ddpoint_pair(dp%).data(0).data0.n(2 * n2% + 1), _
            Ddpoint_pair(dp%).data(0).data0.n(2 * n3%), _
             Ddpoint_pair(dp%).data(0).data0.n(2 * n3% + 1), _
              n_(2), n_(3), l_(1), Ddpoint_pair(dp%).data(0).data0.line_no(n2%), _
              Ddpoint_pair(dp%).data(0).data0.line_no(n3%), _
               l_(1), 1, temp_record, False, 0, 0, 0, 0, False)
       Else
       combine_relation_with_dpoint_pair0_ = set_dpoint_pair( _
         p(2), p(3), Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%), _
           Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1), _
            Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%), _
             Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1), _
              Ddpoint_pair(dp%).data(0).data0.poi(2 * n4%), _
         Ddpoint_pair(dp%).data(0).data0.poi(2 * n4% + 1), _
         n_(2), n_(3), Ddpoint_pair(dp%).data(0).data0.n(2 * n2%), _
           Ddpoint_pair(dp%).data(0).data0.n(2 * n2% + 1), _
            Ddpoint_pair(dp%).data(0).data0.n(2 * n3%), _
             Ddpoint_pair(dp%).data(0).data0.n(2 * n3% + 1), _
              Ddpoint_pair(dp%).data(0).data0.n(2 * n4%), _
         Ddpoint_pair(dp%).data(0).data0.n(2 * n4% + 1), _
          l_(1), Ddpoint_pair(dp%).data(0).data0.line_no(n2%), _
              Ddpoint_pair(dp%).data(0).data0.line_no(n3%), _
         Ddpoint_pair(dp%).data(0).data0.line_no(n4%), _
          1, temp_record, False, 0, 0, 0, 0, False)
        End If
      If combine_relation_with_dpoint_pair0_ > 1 Then
       Exit Function
      End If
      If Ddpoint_pair(dp%).data(0).data0.is_h_ratio = 3 Then
          If Ddpoint_pair(dp%).data(0).data0.poi(0) = p(0) And Ddpoint_pair(dp%).data(0).data0.poi(5) = p(1) Then
           If n1% = 1 Then
            combine_relation_with_dpoint_pair0_ = set_dpoint_pair( _
             Ddpoint_pair(dp%).data(0).data0.poi(0), _
              Ddpoint_pair(dp%).data(0).data0.poi(1), _
               p(2), p(3), p(2), p(3), _
                Ddpoint_pair(dp%).data(0).data0.poi(2), _
                 Ddpoint_pair(dp%).data(0).data0.poi(7), _
             Ddpoint_pair(dp%).data(0).data0.n(0), _
              Ddpoint_pair(dp%).data(0).data0.n(1), _
               n_(2), n_(3), n_(2), n_(3), _
                Ddpoint_pair(dp%).data(0).data0.n(2), _
                 Ddpoint_pair(dp%).data(0).data0.n(7), _
                  Ddpoint_pair(dp%).data(0).data0.line_no(0), _
                   l_(1), l_(1), _
                    Ddpoint_pair(dp%).data(0).data0.line_no(1), _
                     1, temp_record, False, 0, 0, 0, 0, False)
                        If combine_relation_with_dpoint_pair0_ > 1 Then
                         Exit Function
                        End If
            ElseIf n1% = 3 Then
            combine_relation_with_dpoint_pair0_ = set_dpoint_pair( _
             Ddpoint_pair(dp%).data(0).data0.poi(4), _
              Ddpoint_pair(dp%).data(0).data0.poi(5), _
               p(2), p(3), p(2), p(3), _
                Ddpoint_pair(dp%).data(0).data0.poi(2), _
                 Ddpoint_pair(dp%).data(0).data0.poi(7), _
             Ddpoint_pair(dp%).data(0).data0.n(0), _
              Ddpoint_pair(dp%).data(0).data0.n(1), _
               n_(2), n_(3), n_(2), n_(3), _
                Ddpoint_pair(dp%).data(0).data0.n(2), _
                 Ddpoint_pair(dp%).data(0).data0.n(7), _
                  Ddpoint_pair(dp%).data(0).data0.line_no(n2%), _
                   l_(1), l_(1), _
                    Ddpoint_pair(dp%).data(0).data0.line_no(n3%), _
                     1, temp_record, False, 0, 0, 0, 0, False)
                        If combine_relation_with_dpoint_pair0_ > 1 Then
                          Exit Function
                        End If
            End If
          ElseIf Ddpoint_pair(dp%).data(0).data0.poi(2) = p(0) And Ddpoint_pair(dp%).data(0).data0.poi(7) = p(1) Then
           If n1% = 0 Then
            combine_relation_with_dpoint_pair0_ = set_dpoint_pair( _
             Ddpoint_pair(dp%).data(0).data0.poi(0), _
              Ddpoint_pair(dp%).data(0).data0.poi(5), _
               p(2), p(3), p(2), p(3), _
                Ddpoint_pair(dp%).data(0).data0.poi(2), _
                 Ddpoint_pair(dp%).data(0).data0.poi(3), _
             Ddpoint_pair(dp%).data(0).data0.n(0), _
              Ddpoint_pair(dp%).data(0).data0.n(5), _
               n_(2), n_(3), n_(2), n_(3), _
                Ddpoint_pair(dp%).data(0).data0.n(2), _
                 Ddpoint_pair(dp%).data(0).data0.n(3), _
                  Ddpoint_pair(dp%).data(0).data0.line_no(n2%), _
                   l_(1), l_(1), _
                    Ddpoint_pair(dp%).data(0).data0.line_no(n3%), _
                     1, temp_record, False, 0, 0, 0, 0, False)
                      If combine_relation_with_dpoint_pair0_ > 1 Then
                       Exit Function
                      End If
           ElseIf n1% = 2 Then
            combine_relation_with_dpoint_pair0_ = set_dpoint_pair( _
             Ddpoint_pair(dp%).data(0).data0.poi(0), _
              Ddpoint_pair(dp%).data(0).data0.poi(5), _
               p(2), p(3), p(2), p(3), _
                Ddpoint_pair(dp%).data(0).data0.poi(6), _
                 Ddpoint_pair(dp%).data(0).data0.poi(7), _
                  Ddpoint_pair(dp%).data(0).data0.n(0), _
                   Ddpoint_pair(dp%).data(0).data0.n(5), _
                    n_(2), n_(3), n_(2), n_(3), _
             Ddpoint_pair(dp%).data(0).data0.n(6), _
              Ddpoint_pair(dp%).data(0).data0.n(7), _
               Ddpoint_pair(dp%).data(0).data0.line_no(n2%), _
                l_(1), l_(1), _
                 Ddpoint_pair(dp%).data(0).data0.line_no(n3%), _
                   1, temp_record, False, 0, 0, 0, 0, False)
                    If combine_relation_with_dpoint_pair0_ > 1 Then
                     Exit Function
                    End If
           End If
          End If
      ElseIf Ddpoint_pair(dp%).data(0).data0.is_h_ratio = 2 Then
          If Ddpoint_pair(dp%).data(0).data0.poi(0) = p(0) And Ddpoint_pair(dp%).data(0).data0.poi(3) = p(1) Then
           If n1% = 2 Then
            combine_relation_with_dpoint_pair0_ = set_dpoint_pair( _
             Ddpoint_pair(dp%).data(0).data0.poi(0), _
              Ddpoint_pair(dp%).data(0).data0.poi(1), _
               p(2), p(3), p(2), p(3), _
                Ddpoint_pair(dp%).data(0).data0.poi(10), _
                 Ddpoint_pair(dp%).data(0).data0.poi(11), _
             Ddpoint_pair(dp%).data(0).data0.n(0), _
              Ddpoint_pair(dp%).data(0).data0.n(1), _
               n_(2), n_(3), n_(2), n_(3), _
                Ddpoint_pair(dp%).data(0).data0.n(10), _
                 Ddpoint_pair(dp%).data(0).data0.n(11), _
                  Ddpoint_pair(dp%).data(0).data0.line_no(0), _
                   l_(1), l_(1), _
                    Ddpoint_pair(dp%).data(0).data0.line_no(5), _
                     1, temp_record, False, 0, 0, 0, 0, False)
                        If combine_relation_with_dpoint_pair0_ > 1 Then
                         Exit Function
                        End If
            ElseIf n1% = 3 Then
            combine_relation_with_dpoint_pair0_ = set_dpoint_pair( _
             Ddpoint_pair(dp%).data(0).data0.poi(2), _
              Ddpoint_pair(dp%).data(0).data0.poi(3), _
               p(2), p(3), p(2), p(3), _
                Ddpoint_pair(dp%).data(0).data0.poi(10), _
                 Ddpoint_pair(dp%).data(0).data0.poi(11), _
             Ddpoint_pair(dp%).data(0).data0.n(2), _
              Ddpoint_pair(dp%).data(0).data0.n(3), _
               n_(2), n_(3), n_(2), n_(3), _
                Ddpoint_pair(dp%).data(0).data0.n(10), _
                 Ddpoint_pair(dp%).data(0).data0.n(11), _
                  Ddpoint_pair(dp%).data(0).data0.line_no(1), _
                   l_(1), l_(1), _
                    Ddpoint_pair(dp%).data(0).data0.line_no(5), _
                     1, temp_record, False, 0, 0, 0, 0, False)
                        If combine_relation_with_dpoint_pair0_ > 1 Then
                          Exit Function
                        End If
            End If
          ElseIf Ddpoint_pair(dp%).data(0).data0.poi(10) = p(0) And Ddpoint_pair(dp%).data(0).data0.poi(11) = p(1) Then
           If n1% = 0 Then
            combine_relation_with_dpoint_pair0_ = set_dpoint_pair( _
             Ddpoint_pair(dp%).data(0).data0.poi(8), _
              Ddpoint_pair(dp%).data(0).data0.poi(9), _
               p(2), p(3), p(2), p(3), _
                Ddpoint_pair(dp%).data(0).data0.poi(2), _
                 Ddpoint_pair(dp%).data(0).data0.poi(3), _
             Ddpoint_pair(dp%).data(0).data0.n(8), _
              Ddpoint_pair(dp%).data(0).data0.n(9), _
               n_(2), n_(3), n_(2), n_(3), _
                Ddpoint_pair(dp%).data(0).data0.n(2), _
                 Ddpoint_pair(dp%).data(0).data0.n(3), _
                  Ddpoint_pair(dp%).data(0).data0.line_no(4), _
                   l_(1), l_(1), _
                    Ddpoint_pair(dp%).data(0).data0.line_no(1), _
                     1, temp_record, False, 0, 0, 0, 0, False)
                      If combine_relation_with_dpoint_pair0_ > 1 Then
                       Exit Function
                      End If
           ElseIf n1% = 1 Then
            combine_relation_with_dpoint_pair0_ = set_dpoint_pair( _
             Ddpoint_pair(dp%).data(0).data0.poi(8), _
              Ddpoint_pair(dp%).data(0).data0.poi(9), _
               p(2), p(3), p(2), p(3), _
                Ddpoint_pair(dp%).data(0).data0.poi(6), _
                 Ddpoint_pair(dp%).data(0).data0.poi(7), _
                  Ddpoint_pair(dp%).data(0).data0.n(8), _
                   Ddpoint_pair(dp%).data(0).data0.n(9), _
                    n_(2), n_(3), n_(2), n_(3), _
             Ddpoint_pair(dp%).data(0).data0.n(6), _
              Ddpoint_pair(dp%).data(0).data0.n(7), _
               Ddpoint_pair(dp%).data(0).data0.line_no(4), _
                l_(1), l_(1), _
                 Ddpoint_pair(dp%).data(0).data0.line_no(3), _
                   1, temp_record, False, 0, 0, 0, 0, False)
                    If combine_relation_with_dpoint_pair0_ > 1 Then
                     Exit Function
                    End If
            End If
           End If
      End If
    End If
End If
End If
'****************************
If v(1) <> "" Then
 If Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%) = p(4) And _
    Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1) = p(5) Then
   temp_record.record_data = re
      combine_relation_with_dpoint_pair0_ = _
       combine_relation_with_dpoint_pair00_(dp%, _
        n1%, n2%, n3%, n4%, v(1), temp_record.record_data)
   If combine_relation_with_dpoint_pair0_ > 1 Then
    Exit Function
   End If
      combine_relation_with_dpoint_pair0_ = _
       combine_relation_with_h_point_pair0(v(1), Ddpoint_pair(dp%).data(0).data0, _
        n1%, n2%, n3%, n4%, temp_record)
   If combine_relation_with_dpoint_pair0_ > 1 Then
    Exit Function
   End If
 ElseIf Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%) = p(4) And _
    Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1) = p(5) Then
   temp_record.record_data = re
      combine_relation_with_dpoint_pair0_ = _
       combine_relation_with_dpoint_pair00_(dp%, _
        n1%, n3%, n2%, n4%, v(1), temp_record.record_data)
   If combine_relation_with_dpoint_pair0_ > 1 Then
    Exit Function
   End If
       combine_relation_with_dpoint_pair0_ = _
       combine_relation_with_h_point_pair0(v(1), Ddpoint_pair(dp%).data(0).data0, _
        n1%, n3%, n2%, n4%, temp_record)
   If combine_relation_with_dpoint_pair0_ > 1 Then
    Exit Function
   End If
 ElseIf Ddpoint_pair(dp%).data(0).data0.poi(2 * n4%) = p(0) And _
     Ddpoint_pair(dp%).data(0).data0.poi(2 * n4% + 1) = p(1) Then
     temp_record.record_data = re
     If v(1) = "1" Then
      combine_relation_with_dpoint_pair0_ = set_dpoint_pair( _
     p(4), p(5), Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%), _
     Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1), Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%), _
     Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1), p(4), p(5), n_(4), n_(5), _
     Ddpoint_pair(dp%).data(0).data0.n(2 * n2%), Ddpoint_pair(dp%).data(0).data0.n(2 * n2% + 1), _
     Ddpoint_pair(dp%).data(0).data0.n(2 * n3%), Ddpoint_pair(dp%).data(0).data0.n(2 * n3% + 1), _
     n_(4), n_(5), l_(2), Ddpoint_pair(dp%).data(0).data0.line_no(n2%), _
     Ddpoint_pair(dp%).data(0).data0.line_no(n3%), l_(2), 1, _
       temp_record, False, 0, 0, 0, 0, False)
      If combine_relation_with_dpoint_pair0_ > 1 Then
       Exit Function
      End If
     'Else
          If Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%) And _
           Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1) Then
            combine_relation_with_dpoint_pair0_ = set_Drelation( _
             Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%), Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1), _
              p(4), p(5), 0, 0, 0, 0, 0, 0, sqr_string(v(1), True, False), temp_record, 0, 0, 0, 0, 0, False)
      If combine_relation_with_dpoint_pair0_ > 1 Then
       Exit Function
      End If
      Else
      'Call set_item0(p(4), p(5), p(4), p(5), "*", 0, 0, 0, 0, 0, 0, "1", "1", "1,", "", "", 0, _
               condition_data0, 0, it(0), 0)
      'Call set_item0(Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%), Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1), _
       '      Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%), Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1), "*", _
                0, 0, 0, 0, 0, 0, "1", "1", "1,", "", "", 0, _
               condition_data0, 0, it(0), 0)
        'combine_relation_with_dpoint_pair0_ = set_general_string(it(0), it(1), 0, 0, v(1), "-1", "0", "0", "0", 0, _
           0, temp_record, 0, 0)
      'If combine_relation_with_dpoint_pair0_ > 1 Then
       'Exit Function
      'End If
     End If
 End If
  Else
     If v(1) = "1" Then
     temp_record.record_data = re
     combine_relation_with_dpoint_pair0_ = set_dpoint_pair( _
      p(4), p(5), Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%), _
      Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1), Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%), _
      Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1), Ddpoint_pair(dp%).data(0).data0.poi(2 * n4%), _
      Ddpoint_pair(dp%).data(0).data0.poi(2 * n4% + 1), n_(4), n_(5), Ddpoint_pair(dp%).data(0).data0.n(2 * n2%), _
      Ddpoint_pair(dp%).data(0).data0.n(2 * n2% + 1), Ddpoint_pair(dp%).data(0).data0.n(2 * n3%), _
      Ddpoint_pair(dp%).data(0).data0.n(2 * n3% + 1), Ddpoint_pair(dp%).data(0).data0.n(2 * n4%), _
      Ddpoint_pair(dp%).data(0).data0.n(2 * n4% + 1), l_(2), Ddpoint_pair(dp%).data(0).data0.line_no(n2%), _
      Ddpoint_pair(dp%).data(0).data0.line_no(n3%), Ddpoint_pair(dp%).data(0).data0.line_no(n4%), _
       1, temp_record, False, 0, 0, 0, 0, False)
      If combine_relation_with_dpoint_pair0_ > 1 Then
       Exit Function
      End If
     'Else
     ' temp_record.record_data = re
     ' Call set_item0(p(4), p(5), Ddpoint_pair(dp%).data(0).data0.poi(2 * n4%), _
             Ddpoint_pair(dp%).data(0).data0.poi(2 * n4% + 1), "*", n_(4), n_(5), _
              Ddpoint_pair(dp%).data(0).data0.n(2 * n4%), Ddpoint_pair(dp%).data(0).data0.n(2 * n4% + 1), _
               l_(2), Ddpoint_pair(dp%).data(0).data0.line_no(n4%), "1", record_00, "", record_data0, _
                0, it(0), no_reduce)
     ' Call set_item0(Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%), Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1), _
              Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%), Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1), _
               "*", Ddpoint_pair(dp%).data(0).data0.n(2 * n2%), Ddpoint_pair(dp%).data(0).data0.n(2 * n2% + 1), _
                 Ddpoint_pair(dp%).data(0).data0.n(2 * n3%), Ddpoint_pair(dp%).data(0).data0.n(2 * n3% + 1), _
                   Ddpoint_pair(dp%).data(0).data0.line_no(n2%), Ddpoint_pair(dp%).data(0).data0.line_no(n3%), "1", _
                    record_00, "", record_data0, 0, it(1), no_reduce)
     ' combine_relation_with_dpoint_pair0_ = set_general_string(it(0), it(1), 0, 0, _
         v(1), "-1", "0", "0", "0", 0, 0, temp_record, 0, 0)
     ' If combine_relation_with_dpoint_pair0_ > 1 Then
     '  Exit Function
     ' End If
     End If
 End If
End If
'************
End Function

Public Function combine_relation_with_dpoint_pair_( _
  ByVal ty As Byte, ByVal re%, ByVal dp%, k%, l%, no_reduce As Byte) As Byte
'Dim t_n%
Dim temp_record As total_record_type
Dim n(2) As Integer
Dim m(3) As Integer
Dim n_(5) As Integer
Dim p(5) As Integer
Dim l_(2) As Integer
'Dim tl(2) As Integer
'Dim tn(5) As Integer
Dim v(1) As String
Call add_conditions_to_record(dpoint_pair_, dp%, 0, 0, temp_record.record_data.data0.condition_data)
Call add_conditions_to_record(ty, re%, 0, 0, temp_record.record_data.data0.condition_data)
temp_record.record_data.data0.theorem_no = 1
Call read_point_and_ratio_from_relation(ty, re%, k%, p(), n_(), l_(), v(0), v(1))
n(0) = k%
 n(1) = (k% + 1) Mod 3
  n(2) = (k% + 2) Mod 3
If l% = 0 Then
m(0) = 0
 m(1) = 1
  m(2) = 2
   m(3) = 3
ElseIf l% = 1 Then
m(0) = 1
 m(1) = 0
  m(2) = 3
   m(3) = 2
ElseIf l% = 2 Then
m(0) = 2
 m(1) = 3
  m(2) = 0
   m(3) = 1
Else
m(0) = 3
 m(1) = 2
  m(2) = 1
   m(3) = 0
End If
If ty = line_value_ Then
    If m(0) < 4 Then
       combine_relation_with_dpoint_pair_ = _
        combine_line_value_with_dpoint_pair_(re, dp%, m(0), m(1), m(2), m(3), no_reduce)
        If combine_relation_with_dpoint_pair_ > 1 Then
         Exit Function
        End If
    ElseIf l% = 4 Then
       combine_relation_with_dpoint_pair_ = _
        combine_line_value_with_dpoint_pair_(re, dp%, 4, 0, 5, 2, no_reduce)
        If combine_relation_with_dpoint_pair_ > 1 Then
         Exit Function
        End If
       combine_relation_with_dpoint_pair_ = _
        combine_line_value_with_dpoint_pair_(re, dp%, 4, 1, 5, 3, no_reduce)
        If combine_relation_with_dpoint_pair_ > 1 Then
         Exit Function
        End If
    ElseIf l% = 5 Then
       combine_relation_with_dpoint_pair_ = _
        combine_line_value_with_dpoint_pair_(re, dp%, 5, 2, 4, 0, no_reduce)
        If combine_relation_with_dpoint_pair_ > 1 Then
         Exit Function
        End If
       combine_relation_with_dpoint_pair_ = _
        combine_line_value_with_dpoint_pair_(re, dp%, 5, 3, 4, 1, no_reduce)
        If combine_relation_with_dpoint_pair_ > 1 Then
         Exit Function
        End If
    End If
Else
If l% < 4 Then
   combine_relation_with_dpoint_pair_ = _
     combine_relation_with_dpoint_pair0_(p(), n_(), l_(), v(), dp%, _
      m(0), m(1), m(2), m(3), temp_record.record_data, no_reduce)
   If combine_relation_with_dpoint_pair_ > 1 Then
    Exit Function
   End If
ElseIf l% = 4 Then
   combine_relation_with_dpoint_pair_ = _
     combine_relation_with_dpoint_pair0_(p(), n_(), l_(), v(), dp%, _
      4, 0, 5, 2, temp_record.record_data, no_reduce)
   If combine_relation_with_dpoint_pair_ > 1 Then
    Exit Function
   End If
   combine_relation_with_dpoint_pair_ = _
     combine_relation_with_dpoint_pair0_(p(), n_(), l_(), v(), dp%, _
      4, 1, 5, 3, temp_record.record_data, no_reduce)
   If combine_relation_with_dpoint_pair_ > 1 Then
    Exit Function
   End If
ElseIf l% = 5 Then
   combine_relation_with_dpoint_pair_ = _
     combine_relation_with_dpoint_pair0_(p(), n_(), l_(), v(), dp%, _
      5, 2, 4, 0, temp_record.record_data, no_reduce)
   If combine_relation_with_dpoint_pair_ > 1 Then
    Exit Function
   End If
   combine_relation_with_dpoint_pair_ = _
     combine_relation_with_dpoint_pair0_(p(), n_(), l_(), v(), dp%, _
      5, 3, 4, 1, temp_record.record_data, no_reduce)
   If combine_relation_with_dpoint_pair_ > 1 Then
    Exit Function
   End If
End If
End If
End Function

Public Function combine_dpoint_pair_with_relation(ByVal dp%, _
                     ByVal start%, ByVal no_reduce As Byte) As Byte '10.10
Dim i%, j%, k%, no%
Dim v(1) As String
Dim n_(1) As Integer
Dim n(3) As Integer
Dim m(2) As Integer
Dim tn() As Integer
Dim last_tn%
Dim rel As relation_data0_type
'On Error GoTo combine_dpoint_pair_with_relation_error
If Ddpoint_pair(dp%).record_.no_reduce > 4 Then
 Exit Function
End If
For i% = 0 To 3
n(0) = i%
n(1) = (i% + 1) Mod 4
n(2) = (i% + 2) Mod 4
n(3) = (i% + 3) Mod 4
For j% = 0 To 2
m(0) = j%
m(1) = (j% + 1) Mod 3
m(2) = (j% + 2) Mod 3
rel.poi(2 * m(0)) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n(0))
rel.poi(2 * m(0) + 1) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n(0) + 1)
rel.poi(2 * m(1)) = -1
Call search_for_relation(rel, m(0), n_(0), 1)
rel.poi(2 * m(1)) = 30000
Call search_for_relation(rel, m(0), n_(1), 1)  '5.7
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
no% = Drelation(k%).data(0).record.data1.index.i(m(0))
If no% > start% And Drelation(no%).record_.no_reduce < 4 Then
' If is_two_record_related(relation_, no%, Drelation(no%).data(0).record, _
    dpoint_pair_, dp%, Ddpoint_pair(dp%).data(0).record) = False Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
'End If
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
combine_dpoint_pair_with_relation = _
 combine_relation_with_dpoint_pair_(relation_, no%, dp%, m(0), n(0), no_reduce)
If combine_dpoint_pair_with_relation > 1 Then
 Exit Function
End If
Next k%
Next j%
Next i%
Exit Function
combine_dpoint_pair_with_relation_error:
combine_dpoint_pair_with_relation = 0
End Function
Public Function combine_dpoint_pair_with_dpoint_pair_(ByVal dp1%, _
            ByVal n1%, ByVal n2%, ByVal n3%, ByVal n4%, ByVal dp2%, _
             ByVal m1%, ByVal m2%, ByVal m3%, ByVal m4%, _
              ByVal no_reduce As Byte) As Byte
Dim temp_record As total_record_type
Dim dp(1) As point_pair_data0_type
Dim dp_(1) As point_pair_data0_type
Dim con_ty As Byte
Dim tn%
If dp1% > dp2% Then
 tn% = dp2%
ElseIf dp1% < dp2% Then
 tn% = dp1%
Else
 Exit Function
End If
Call add_conditions_to_record(dpoint_pair_, dp1%, dp2%, 0, temp_record.record_data.data0.condition_data)
     temp_record.record_data.data0.theorem_no = 1
dp(0).poi(0) = Ddpoint_pair(dp1%).data(0).data0.poi(2 * n1%)
dp(0).poi(1) = Ddpoint_pair(dp1%).data(0).data0.poi(2 * n1% + 1)
dp(0).poi(2) = Ddpoint_pair(dp1%).data(0).data0.poi(2 * n2%)
dp(0).poi(3) = Ddpoint_pair(dp1%).data(0).data0.poi(2 * n2% + 1)
dp(0).poi(4) = Ddpoint_pair(dp1%).data(0).data0.poi(2 * n3%)
dp(0).poi(5) = Ddpoint_pair(dp1%).data(0).data0.poi(2 * n3% + 1)
dp(0).poi(6) = Ddpoint_pair(dp1%).data(0).data0.poi(2 * n4%)
dp(0).poi(7) = Ddpoint_pair(dp1%).data(0).data0.poi(2 * n4% + 1)
dp(0).n(0) = Ddpoint_pair(dp1%).data(0).data0.n(2 * n1%)
dp(0).n(1) = Ddpoint_pair(dp1%).data(0).data0.n(2 * n1% + 1)
dp(0).n(2) = Ddpoint_pair(dp1%).data(0).data0.n(2 * n2%)
dp(0).n(3) = Ddpoint_pair(dp1%).data(0).data0.n(2 * n2% + 1)
dp(0).n(4) = Ddpoint_pair(dp1%).data(0).data0.n(2 * n3%)
dp(0).n(5) = Ddpoint_pair(dp1%).data(0).data0.n(2 * n3% + 1)
dp(0).n(6) = Ddpoint_pair(dp1%).data(0).data0.n(2 * n4%)
dp(0).n(7) = Ddpoint_pair(dp1%).data(0).data0.n(2 * n4% + 1)
dp(0).line_no(0) = Ddpoint_pair(dp1%).data(0).data0.line_no(n1%)
dp(0).line_no(1) = Ddpoint_pair(dp1%).data(0).data0.line_no(n2%)
dp(0).line_no(2) = Ddpoint_pair(dp1%).data(0).data0.line_no(n3%)
dp(0).line_no(3) = Ddpoint_pair(dp1%).data(0).data0.line_no(n4%)
If (n1% = 0 And n2% = 1) Or (n1% = 1 And n2% = 0) Then
 dp(0).poi(8) = Ddpoint_pair(dp1%).data(0).data0.poi(8)
 dp(0).poi(9) = Ddpoint_pair(dp1%).data(0).data0.poi(9)
 dp(0).n(8) = Ddpoint_pair(dp1%).data(0).data0.n(8)
 dp(0).n(9) = Ddpoint_pair(dp1%).data(0).data0.n(9)
 dp(0).line_no(4) = Ddpoint_pair(dp1%).data(0).data0.line_no(4)
 dp(0).con_line_type(0) = Ddpoint_pair(dp1%).data(0).data0.con_line_type(0)
 dp(0).poi(10) = Ddpoint_pair(dp1%).data(0).data0.poi(10)
 dp(0).poi(11) = Ddpoint_pair(dp1%).data(0).data0.poi(11)
 dp(0).n(10) = Ddpoint_pair(dp1%).data(0).data0.n(10)
 dp(0).n(11) = Ddpoint_pair(dp1%).data(0).data0.n(11)
 dp(0).line_no(5) = Ddpoint_pair(dp1%).data(0).data0.line_no(5)
 dp(0).con_line_type(1) = Ddpoint_pair(dp1%).data(0).data0.con_line_type(1)
ElseIf (n1% = 2 And n2% = 3) Or (n1% = 3 And n2% = 2) Then
 dp(0).poi(8) = Ddpoint_pair(dp1%).data(0).data0.poi(10)
 dp(0).poi(9) = Ddpoint_pair(dp1%).data(0).data0.poi(11)
 dp(0).n(8) = Ddpoint_pair(dp1%).data(0).data0.n(10)
 dp(0).n(9) = Ddpoint_pair(dp1%).data(0).data0.n(11)
 dp(0).line_no(4) = Ddpoint_pair(dp1%).data(0).data0.line_no(5)
 dp(0).con_line_type(0) = Ddpoint_pair(dp1%).data(0).data0.con_line_type(1)
 dp(0).poi(10) = Ddpoint_pair(dp1%).data(0).data0.poi(8)
 dp(0).poi(11) = Ddpoint_pair(dp1%).data(0).data0.poi(9)
 dp(0).n(10) = Ddpoint_pair(dp1%).data(0).data0.n(8)
 dp(0).n(11) = Ddpoint_pair(dp1%).data(0).data0.n(9)
 dp(0).line_no(5) = Ddpoint_pair(dp1%).data(0).data0.line_no(4)
 dp(0).con_line_type(1) = Ddpoint_pair(dp1%).data(0).data0.con_line_type(0)
End If
dp(1).poi(0) = Ddpoint_pair(dp2%).data(0).data0.poi(2 * m1%)
dp(1).poi(1) = Ddpoint_pair(dp2%).data(0).data0.poi(2 * m1% + 1)
dp(1).poi(2) = Ddpoint_pair(dp2%).data(0).data0.poi(2 * m2%)
dp(1).poi(3) = Ddpoint_pair(dp2%).data(0).data0.poi(2 * m2% + 1)
dp(1).poi(4) = Ddpoint_pair(dp2%).data(0).data0.poi(2 * m3%)
dp(1).poi(5) = Ddpoint_pair(dp2%).data(0).data0.poi(2 * m3% + 1)
dp(1).poi(6) = Ddpoint_pair(dp2%).data(0).data0.poi(2 * m4%)
dp(1).poi(7) = Ddpoint_pair(dp2%).data(0).data0.poi(2 * m4% + 1)
dp(1).n(0) = Ddpoint_pair(dp2%).data(0).data0.n(2 * m1%)
dp(1).n(1) = Ddpoint_pair(dp2%).data(0).data0.n(2 * m1% + 1)
dp(1).n(2) = Ddpoint_pair(dp2%).data(0).data0.n(2 * m2%)
dp(1).n(3) = Ddpoint_pair(dp2%).data(0).data0.n(2 * m2% + 1)
dp(1).n(4) = Ddpoint_pair(dp2%).data(0).data0.n(2 * m3%)
dp(1).n(5) = Ddpoint_pair(dp2%).data(0).data0.n(2 * m3% + 1)
dp(1).n(6) = Ddpoint_pair(dp2%).data(0).data0.n(2 * m4%)
dp(1).n(7) = Ddpoint_pair(dp2%).data(0).data0.n(2 * m4% + 1)
dp(1).line_no(0) = Ddpoint_pair(dp2%).data(0).data0.line_no(m1%)
dp(1).line_no(1) = Ddpoint_pair(dp2%).data(0).data0.line_no(m2%)
dp(1).line_no(2) = Ddpoint_pair(dp2%).data(0).data0.line_no(m3%)
dp(1).line_no(3) = Ddpoint_pair(dp2%).data(0).data0.line_no(m4%)
If (m1% = 0 And m2% = 1) Or (m1% = 1 And m2% = 0) Then
 dp(1).poi(8) = Ddpoint_pair(dp2%).data(0).data0.poi(8)
 dp(1).poi(9) = Ddpoint_pair(dp2%).data(0).data0.poi(9)
 dp(1).n(8) = Ddpoint_pair(dp2%).data(0).data0.n(8)
 dp(1).n(9) = Ddpoint_pair(dp2%).data(0).data0.n(9)
 dp(1).line_no(4) = Ddpoint_pair(dp2%).data(0).data0.line_no(4)
 dp(1).con_line_type(0) = Ddpoint_pair(dp2%).data(0).data0.con_line_type(0)
 dp(1).poi(10) = Ddpoint_pair(dp2%).data(0).data0.poi(10)
 dp(1).poi(11) = Ddpoint_pair(dp2%).data(0).data0.poi(11)
 dp(1).n(10) = Ddpoint_pair(dp2%).data(0).data0.n(10)
 dp(1).n(11) = Ddpoint_pair(dp2%).data(0).data0.n(11)
 dp(1).line_no(5) = Ddpoint_pair(dp2%).data(0).data0.line_no(5)
 dp(1).con_line_type(1) = Ddpoint_pair(dp2%).data(0).data0.con_line_type(1)
ElseIf (m1% = 2 And m2% = 3) Or (m1% = 3 And m2% = 2) Then
 dp(1).poi(8) = Ddpoint_pair(dp2%).data(0).data0.poi(10)
 dp(1).poi(9) = Ddpoint_pair(dp2%).data(0).data0.poi(11)
 dp(1).n(8) = Ddpoint_pair(dp2%).data(0).data0.n(10)
 dp(1).n(9) = Ddpoint_pair(dp2%).data(0).data0.n(11)
 dp(1).line_no(4) = Ddpoint_pair(dp2%).data(0).data0.line_no(5)
 dp(1).con_line_type(0) = Ddpoint_pair(dp2%).data(0).data0.con_line_type(1)
 dp(1).poi(10) = Ddpoint_pair(dp2%).data(0).data0.poi(8)
 dp(1).poi(11) = Ddpoint_pair(dp2%).data(0).data0.poi(9)
 dp(1).n(10) = Ddpoint_pair(dp2%).data(0).data0.n(8)
 dp(1).n(11) = Ddpoint_pair(dp2%).data(0).data0.n(9)
 dp(1).line_no(5) = Ddpoint_pair(dp2%).data(0).data0.line_no(4)
 dp(1).con_line_type(1) = Ddpoint_pair(dp2%).data(0).data0.con_line_type(0)
End If
dp_(0) = dp(0)
dp_(1) = dp(1)
combine_dpoint_pair_with_dpoint_pair_ = _
 combine_dpoint_pair_with_dpoint_pair0(dp_(0), dp_(1), temp_record.record_data)
If combine_dpoint_pair_with_dpoint_pair_ > 1 Then
 Exit Function
End If
If (dp(0).con_line_type(0) = 3 Or _
     dp(0).con_line_type(0) = 5) And _
      (dp(0).con_line_type(1) = 3 Or _
        dp(0).con_line_type(1) = 5) Then
 dp_(0) = dp(0)
 dp_(0).poi(2) = dp(0).poi(8)
 dp_(0).poi(3) = dp(0).poi(9)
 dp_(0).n(2) = dp(0).n(8)
 dp_(0).n(3) = dp(0).n(9)
 dp_(0).line_no(1) = dp(0).line_no(4)
 dp_(0).poi(6) = dp(0).poi(10)
 dp_(0).poi(7) = dp(0).poi(11)
 dp_(0).n(6) = dp(0).n(10)
 dp_(0).n(7) = dp(0).n(11)
 dp_(0).line_no(3) = dp(0).line_no(5)
 dp_(1) = dp(1)
combine_dpoint_pair_with_dpoint_pair_ = _
 combine_dpoint_pair_with_dpoint_pair0(dp_(0), dp_(1), temp_record.record_data)
If combine_dpoint_pair_with_dpoint_pair_ > 1 Then
 Exit Function
End If
End If
If (dp(1).con_line_type(0) = 3 And _
     dp(1).con_line_type(0) = 5) And _
        (Ddpoint_pair(dp2%).data(0).data0.con_line_type(1) = 3 Or _
          Ddpoint_pair(dp2%).data(0).data0.con_line_type(1) = 5) Then
 dp_(1) = dp(1)
 dp_(1).poi(2) = dp(1).poi(8)
 dp_(1).poi(3) = dp(1).poi(9)
 dp_(1).n(2) = dp(1).n(8)
 dp_(1).n(3) = dp(1).n(9)
 dp_(1).line_no(1) = dp(1).line_no(4)
 dp_(1).poi(6) = dp(1).poi(10)
 dp_(1).poi(7) = dp(1).poi(11)
 dp_(1).n(6) = dp(1).n(10)
 dp_(1).n(7) = dp(1).n(11)
 dp_(1).line_no(3) = dp(1).line_no(5)
dp_(0) = dp(0)
combine_dpoint_pair_with_dpoint_pair_ = _
 combine_dpoint_pair_with_dpoint_pair0(dp_(0), dp_(1), temp_record.record_data)
If combine_dpoint_pair_with_dpoint_pair_ > 1 Then
 Exit Function
End If
End If
End Function
Public Function combine_dpoint_pair_with_dpoint_pair0( _
     dp1 As point_pair_data0_type, dp2 As point_pair_data0_type, _
         re As record_data_type) As Byte '10.10
Dim temp_record As total_record_type
Dim dn(2) As Integer
Dim con_ty As Byte
Dim tn%
re.data0.theorem_no = 1
If is_equal_dline(dp1.poi(2), dp1.poi(3), _
     dp2.poi(2), dp2.poi(3), dp1.n(2), dp1.n(3), _
      dp2.n(2), dp2.n(3), dp1.line_no(1), dp2.line_no(1), _
       dn(0), -1000, 0, 0, 0, eline_data0, dn(1), dn(2), _
        con_ty, "", record_0.data0.condition_data) Then
  temp_record.record_data = re
  Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), temp_record.record_data.data0.condition_data)
  record_0.data0.condition_data.condition_no = 0 'record0
     If is_equal_dline(dp1.poi(4), dp1.poi(5), _
      dp2.poi(4), dp2.poi(5), dp1.n(4), dp1.n(5), _
       dp2.n(4), dp2.n(5), dp1.line_no(2), dp2.line_no(2), _
        dn(0), -1000, 0, 0, 0, eline_data0, dn(1), dn(2), _
         con_ty, "", record_0.data0.condition_data) Then
     Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), temp_record.record_data.data0.condition_data)
      combine_dpoint_pair_with_dpoint_pair0 = set_equal_dline( _
       dp1.poi(6), dp1.poi(7), dp2.poi(6), dp2.poi(7), _
        dp1.n(6), dp1.n(7), dp2.n(6), dp2.n(7), _
         dp1.line_no(3), dp2.line_no(3), 0, temp_record, 0, 0, 0, 0, 0, False)
          If combine_dpoint_pair_with_dpoint_pair0 > 1 Then
           Exit Function
          End If
      ElseIf is_equal_dline(dp1.poi(6), dp1.poi(7), _
            dp2.poi(6), dp2.poi(7), dp1.n(6), dp1.n(7), _
             dp2.n(6), dp2.n(7), dp1.line_no(3), dp2.line_no(3), _
              dn(0), -1000, 0, 0, 0, eline_data0, dn(1), dn(2), _
               con_ty, "", record_0.data0.condition_data) Then
         Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), temp_record.record_data.data0.condition_data)
          combine_dpoint_pair_with_dpoint_pair0 = set_equal_dline( _
           dp1.poi(4), dp1.poi(5), dp2.poi(4), dp2.poi(5), _
            dp1.n(4), dp1.n(5), dp2.n(4), dp2.n(5), _
             dp1.line_no(2), dp2.line_no(2), 0, temp_record, 0, 0, 0, 0, 0, False)
              If combine_dpoint_pair_with_dpoint_pair0 > 1 Then
               Exit Function
              End If
      Else
       combine_dpoint_pair_with_dpoint_pair0 = set_dpoint_pair( _
         dp1.poi(4), dp1.poi(5), dp1.poi(6), dp1.poi(7), _
          dp2.poi(4), dp2.poi(5), dp2.poi(6), dp2.poi(7), _
           dp1.n(4), dp1.n(5), dp1.n(6), dp1.n(7), _
            dp2.n(4), dp2.n(5), dp2.n(6), dp2.n(7), _
             dp1.line_no(2), dp1.line_no(3), dp2.line_no(2), dp2.line_no(3), _
              0, temp_record, True, 0, 0, 0, 0, False)
               If combine_dpoint_pair_with_dpoint_pair0 > 1 Then
                Exit Function
               End If
      End If
 ElseIf is_equal_dline(dp1.poi(2), dp1.poi(3), _
           dp2.poi(4), dp2.poi(5), dp1.n(2), dp1.n(3), _
            dp2.n(4), dp2.n(5), dp1.line_no(1), dp2.line_no(2), _
             dn(0), -1000, 0, 0, 0, eline_data0, dn(1), dn(2), _
              con_ty, "", record_0.data0.condition_data) Then
     temp_record.record_data = re
    Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), temp_record.record_data.data0.condition_data)
     If is_equal_dline(dp1.poi(4), dp1.poi(5), _
            dp2.poi(2), dp2.poi(3), dp1.n(4), dp1.n(5), _
             dp2.n(2), dp2.n(3), dp1.line_no(2), dp2.line_no(1), _
              dn(0), -1000, 0, 0, 0, eline_data0, dn(1), dn(2), _
               con_ty, "", record_0.data0.condition_data) Then
      Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), temp_record.record_data.data0.condition_data)
       combine_dpoint_pair_with_dpoint_pair0 = set_equal_dline( _
        dp1.poi(6), dp1.poi(7), dp2.poi(6), dp2.poi(7), _
         dp1.n(6), dp1.n(7), dp2.n(6), dp2.n(7), _
          dp1.line_no(3), dp2.line_no(3), 0, temp_record, 0, 0, 0, 0, 0, False)
           If combine_dpoint_pair_with_dpoint_pair0 > 1 Then
            Exit Function
           End If
     ElseIf is_equal_dline(dp1.poi(6), dp1.poi(7), _
            dp2.poi(6), dp2.poi(7), dp1.n(6), dp1.n(7), _
             dp2.n(6), dp2.n(7), dp1.line_no(3), dp2.line_no(3), _
              dn(0), -1000, 0, 0, 0, eline_data0, dn(1), dn(2), _
               con_ty, "", record_0.data0.condition_data) Then
       Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), temp_record.record_data.data0.condition_data)
        combine_dpoint_pair_with_dpoint_pair0 = set_equal_dline( _
         dp1.poi(4), dp1.poi(5), dp2.poi(2), dp2.poi(3), _
          dp1.n(4), dp1.n(5), dp2.n(2), dp2.n(3), _
           dp1.line_no(2), dp2.line_no(1), 0, temp_record, 0, 0, 0, 0, 0, False)
            If combine_dpoint_pair_with_dpoint_pair0 > 1 Then
             Exit Function
            End If
    Else
     combine_dpoint_pair_with_dpoint_pair0 = set_dpoint_pair( _
      dp1.poi(4), dp1.poi(5), dp1.poi(6), dp1.poi(7), _
       dp2.poi(2), dp2.poi(3), dp2.poi(6), dp2.poi(7), _
        dp1.n(4), dp1.n(5), dp1.n(6), dp1.n(7), _
         dp2.n(2), dp2.n(3), dp2.n(6), dp2.n(7), _
          dp1.line_no(2), dp1.line_no(3), dp2.line_no(1), dp2.line_no(3), _
           0, temp_record, True, 0, 0, 0, 0, False)
            If combine_dpoint_pair_with_dpoint_pair0 > 1 Then
             Exit Function
            End If
   End If
 ElseIf is_equal_dline(dp1.poi(4), dp1.poi(5), _
          dp2.poi(2), dp2.poi(3), dp1.n(4), dp1.n(5), _
           dp2.n(2), dp2.n(3), dp1.line_no(2), dp2.line_no(1), _
            dn(0), -1000, 0, 0, 0, eline_data0, _
             dn(1), dn(2), con_ty, "", record_0.data0.condition_data) Then
  temp_record.record_data = re
  Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), temp_record.record_data.data0.condition_data)
   If is_equal_dline(dp1.poi(6), dp1.poi(7), _
          dp2.poi(6), dp2.poi(7), dp1.n(6), dp1.n(7), _
           dp2.n(6), dp2.n(7), dp1.line_no(3), dp2.line_no(3), _
            dn(0), -1000, 0, 0, 0, eline_data0, dn(1), dn(2), _
             con_ty, "", record_0.data0.condition_data) Then
    Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), temp_record.record_data.data0.condition_data)
     combine_dpoint_pair_with_dpoint_pair0 = set_equal_dline( _
      dp1.poi(2), dp1.poi(3), dp2.poi(4), dp2.poi(5), _
       dp1.n(2), dp1.n(3), dp2.n(4), dp2.n(5), _
        dp1.line_no(1), dp2.line_no(2), 0, temp_record, 0, 0, 0, 0, 0, False)
         If combine_dpoint_pair_with_dpoint_pair0 > 1 Then
          Exit Function
         End If
   Else
    combine_dpoint_pair_with_dpoint_pair0 = set_dpoint_pair( _
       dp1.poi(2), dp1.poi(3), dp1.poi(6), dp1.poi(7), _
        dp2.poi(4), dp2.poi(5), dp2.poi(6), dp2.poi(7), _
         dp1.n(2), dp1.n(3), dp1.n(6), dp1.n(7), _
          dp2.n(4), dp2.n(5), dp2.n(6), dp2.n(7), _
           dp1.line_no(1), dp1.line_no(3), dp2.line_no(2), dp2.line_no(3), _
            0, temp_record, True, 0, 0, 0, 0, False)
             If combine_dpoint_pair_with_dpoint_pair0 > 1 Then
              Exit Function
             End If
   End If
ElseIf is_equal_dline(dp1.poi(4), dp1.poi(5), _
        dp2.poi(4), dp2.poi(5), dp1.n(4), dp1.n(5), _
         dp2.n(4), dp2.n(5), dp1.line_no(2), dp2.line_no(2), _
          dn(0), -1000, 0, 0, 0, eline_data0, _
           dn(1), dn(2), con_ty, "", record_0.data0.condition_data) Then
   temp_record.record_data = re
    Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), temp_record.record_data.data0.condition_data)
 If is_equal_dline(dp1.poi(6), dp1.poi(7), _
            dp2.poi(6), dp2.poi(7), dp1.n(6), dp1.n(7), _
             dp2.n(6), dp2.n(7), dp1.line_no(3), dp2.line_no(3), _
              dn(0), -1000, 0, 0, 0, eline_data0, _
               dn(1), dn(2), con_ty, "", record_0.data0.condition_data) Then
  Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), temp_record.record_data.data0.condition_data)
    combine_dpoint_pair_with_dpoint_pair0 = set_equal_dline( _
     dp1.poi(2), dp1.poi(3), dp2.poi(2), dp2.poi(2), _
      dp1.n(2), dp1.n(3), dp2.n(2), dp2.n(3), _
       dp1.line_no(1), dp2.line_no(1), 0, temp_record, 0, 0, 0, 0, 0, False)
        If combine_dpoint_pair_with_dpoint_pair0 > 1 Then
         Exit Function
        End If
  Else
      combine_dpoint_pair_with_dpoint_pair0 = set_dpoint_pair( _
       dp1.poi(2), dp1.poi(3), dp1.poi(6), dp1.poi(7), _
        dp2.poi(2), dp2.poi(3), dp2.poi(6), dp2.poi(7), _
         dp1.n(2), dp1.n(3), dp1.n(6), dp1.n(7), _
          dp2.n(2), dp2.n(3), dp2.n(6), dp2.n(7), _
           dp1.line_no(1), dp1.line_no(3), dp2.line_no(1), dp2.line_no(3), _
            0, temp_record, True, 0, 0, 0, 0, False)
             If combine_dpoint_pair_with_dpoint_pair0 > 1 Then
              Exit Function
             End If
  End If
ElseIf is_equal_dline(dp1.poi(6), dp1.poi(7), _
     dp2.poi(6), dp2.poi(7), dp1.n(6), dp1.n(7), _
      dp2.n(6), dp2.n(7), dp1.line_no(3), dp2.line_no(3), _
       dn(0), -1000, 0, 0, 0, eline_data0, dn(1), dn(2), _
        con_ty, "", record_0.data0.condition_data) Then
 temp_record.record_data = re
  Call add_conditions_to_record(con_ty, dn(0), dn(1), dn(2), temp_record.record_data.data0.condition_data)
  combine_dpoint_pair_with_dpoint_pair0 = set_dpoint_pair( _
       dp1.poi(2), dp1.poi(3), dp2.poi(2), dp2.poi(3), _
        dp2.poi(4), dp2.poi(5), dp1.poi(4), dp1.poi(5), _
         dp1.n(2), dp1.n(3), dp2.n(2), dp2.n(3), _
          dp2.n(4), dp2.n(5), dp1.n(4), dp1.n(5), _
           dp1.line_no(1), dp2.line_no(1), dp2.line_no(2), dp1.line_no(2), _
            0, temp_record, True, 0, 0, 0, 0, False)
End If
End Function
Public Function combine_two_point(ByVal p1%, ByVal p2%, no%, re As total_record_type) As Byte
' 两点重合
Dim i%
Dim temp_record(1) As total_record_type
If p1% > p2% Then
 Call exchange_two_integer(p1%, p2%)
End If
For i% = 1 To last_conditions.last_cond(1).two_point_conset_no
 If two_point_conset(i%).data(0).poi(0) = p1% And _
      two_point_conset(i%).data(0).poi(1) = p2% Then
       no% = i%
        Exit Function
 End If
Next i%
If last_conditions.last_cond(1).two_point_conset_no Mod 10 = 0 Then
ReDim Preserve two_point_conset(last_conditions.last_cond(1).two_point_conset_no + 10) _
         As two_point_conset_type
End If
last_conditions.last_cond(1).two_point_conset_no = _
   last_conditions.last_cond(1).two_point_conset_no + 1
    no% = last_conditions.last_cond(1).two_point_conset_no
two_point_conset(no%).data(0).poi(0) = p1%
two_point_conset(no%).data(0).poi(0) = p2%
two_point_conset(no%).data(0).record = re.record_data
two_point_conset(no%).record_ = re.record_
temp_record(0).record_data.data0.condition_data.condition_no = 1
temp_record(0).record_data.data0.condition_data.condition(1).ty = two_point_conset_
temp_record(0).record_data.data0.condition_data.condition(1).no = no%
End Function


Public Function combine_two_tangent_line_(ByVal tl1%, ByVal tl2%, no_reduce As Byte) As Byte
Dim temp_record As total_record_type
Dim p%, n1%, n2%
temp_record.record_data.data0.condition_data.condition_no = 2
temp_record.record_data.data0.condition_data.condition(0).ty = tangent_line_
temp_record.record_data.data0.condition_data.condition(0).no = tl1%
temp_record.record_data.data0.condition_data.condition(1).ty = tangent_line_
temp_record.record_data.data0.condition_data.condition(1).no = tl2%
If is_same_two_point(tangent_line(tl1%).data(0).circ(0), tangent_line(tl2%).data(0).circ(0), _
    tangent_line(tl1%).data(0).circ(1), tangent_line(tl2%).data(0).circ(1)) And _
     tangent_line(tl1%).data(0).circ(1) > 0 Then
 combine_two_tangent_line_ = set_equal_dline(tangent_line(tl1%).data(0).poi(0), _
         tangent_line(tl1%).data(0).poi(1), tangent_line(tl2%).data(0).poi(0), _
          tangent_line(tl2%).data(0).poi(1), tangent_line(tl1%).data(0).n(0), _
           tangent_line(tl1%).data(0).n(1), tangent_line(tl2%).data(0).n(0), _
            tangent_line(tl2%).data(0).n(1), tangent_line(tl1%).data(0).line_no, _
             tangent_line(tl2%).data(0).line_no, 0, temp_record, 0, 0, 0, 0, no_reduce, False)
 If combine_two_tangent_line_ > 1 Then
  Exit Function
 End If
Else
 p% = is_line_line_intersect(tangent_line(tl1%).data(0).line_no, _
             tangent_line(tl2%).data(0).line_no, n1%, n2%, False)
 If p% > 0 Then
  If tangent_line(tl1%).data(0).circ(0) = tangent_line(tl2%).data(0).circ(0) Then
    combine_two_tangent_line_ = set_equal_dline(tangent_line(tl1%).data(0).poi(0), _
         p%, tangent_line(tl2%).data(0).poi(0), p%, tangent_line(tl1%).data(0).n(0), _
          n1%, tangent_line(tl2%).data(0).n(0), n2%, tangent_line(tl1%).data(0).line_no, _
           tangent_line(tl2%).data(0).line_no, 0, temp_record, 0, 0, 0, 0, no_reduce, False)
   If combine_two_tangent_line_ > 1 Then
    Exit Function
   End If
  ElseIf tangent_line(tl1%).data(0).circ(0) = tangent_line(tl2%).data(0).circ(1) Then
    combine_two_tangent_line_ = set_equal_dline(tangent_line(tl1%).data(0).poi(0), _
         p%, tangent_line(tl2%).data(0).poi(1), p%, tangent_line(tl1%).data(0).n(0), _
          n1%, tangent_line(tl2%).data(0).n(1), n2%, tangent_line(tl1%).data(0).line_no, _
           tangent_line(tl2%).data(0).line_no, 0, temp_record, 0, 0, 0, 0, no_reduce, False)
   If combine_two_tangent_line_ > 1 Then
    Exit Function
   End If
  ElseIf tangent_line(tl1%).data(0).circ(1) = tangent_line(tl2%).data(0).circ(0) Then
    combine_two_tangent_line_ = set_equal_dline(tangent_line(tl1%).data(0).poi(1), _
         p%, tangent_line(tl2%).data(0).poi(0), p%, tangent_line(tl1%).data(0).n(1), _
          n1%, tangent_line(tl2%).data(0).n(0), n2%, tangent_line(tl1%).data(0).line_no, _
           tangent_line(tl2%).data(0).line_no, 0, temp_record, 0, 0, 0, 0, no_reduce, False)
   If combine_two_tangent_line_ > 1 Then
    Exit Function
   End If
  End If
 End If
End If
End Function
Public Function combine_tangent_line_with_tangent_line(ByVal tl%, no_reduce As Byte) As Byte
Dim i%, k%
For k% = 1 + last_conditions.last_cond(0).tangent_line_no To last_conditions.last_cond(1).tangent_line_no
 i% = tangent_line(k%).data(0).record.data1.index.i(0)
 If i% > tl% Then
  combine_tangent_line_with_tangent_line = combine_two_tangent_line_( _
    i%, tl%, no_reduce)
  If combine_tangent_line_with_tangent_line > 1 Then
   Exit Function
  End If
 End If
Next k%
End Function

Public Function combine_six_angle_(ByVal A1%, ByVal A2%, ByVal A3%, _
        ByVal A4%, ByVal A5%, ByVal A6%, ByVal p1$, ByVal p2$, _
         ByVal p3$, ByVal p4$, ByVal p5$, ByVal p6$, v$, re As record_data_type) As Byte
Dim A(5) As Integer
Dim p(5) As String
Dim i%, j%, k%
Dim temp_record As total_record_type
A(0) = A1%
A(1) = A2%
A(2) = A3%
A(3) = A4%
A(4) = A5%
A(5) = A6%
p(0) = p1$
p(1) = p2$
p(2) = p3$
p(3) = p4$
p(4) = p5$
p(5) = p6$
For i% = 0 To 4
 For j% = i% + 1 To 5
  If A(i%) > 0 Then
  If A(i%) = A(j%) Then
   p(i%) = add_string(p(i%), p(j%), True, False)
    p(j%) = "0"
     A(j%) = "0"
   If p(i%) = "0" Then
    A(i%) = 0
   End If
  End If
  End If
 Next j%
Next i%
Call remove_record_for_zero_para(p(), A(), 5)
If A(3) = 0 Then
temp_record.record_data = re
combine_six_angle_ = set_three_angle_value(A(0), A(1), A(2), _
 p(0), p(1), p(2), v$, 0, temp_record, 0, 0, 0, 0, 0, 0, False)
 If combine_six_angle_ > 1 Then
  Exit Function
 End If
End If
End Function
Public Function combine_six_angle0(ByVal A1%, ByVal A2%, ByVal A3%, _
        ByVal A4%, ByVal A5%, ByVal A6%, ByVal p1$, ByVal p2$, _
         ByVal p3$, ByVal p4$, ByVal p5$, ByVal p6$, ByVal v$, last_angle%, _
           re As record_data_type) As Byte
Dim A(5) As Integer
Dim p(5) As String
Dim tA  As Integer
Dim i%, j%, k%
Dim temp_record As total_record_type
A(0) = A1%
A(1) = A2%
A(2) = A3%
A(3) = A4%
A(4) = A5%
A(5) = A6%
p(0) = p1$
p(1) = p2$
p(2) = p3$
p(3) = p4$
p(4) = p5$
p(5) = p6$
temp_record.record_data = re
For i% = 0 To last_angle% - 2
 For j% = i% + 1 To last_angle% - 1
  If A(i%) > 0 Then
  If A(i%) = A(j%) Then
   p(i%) = add_string(p(i%), p(j%), True, False)
    p(j%) = "0"
     A(j%) = "0"
   If p(i%) = "0" Then
    A(i%) = 0
   End If
  End If
  End If
 Next j%
Next i%
Call remove_record_for_zero_para(p(), A(), last_angle%)
If last_angle% <= 2 Then
combine_six_angle0 = set_three_angle_value(A(0), A(1), A(2), _
 p(0), p(1), p(2), v$, 0, temp_record, 0, 0, 0, 0, 0, 0, False)
 If combine_six_angle0 > 1 Then
  Exit Function
 End If
Else
 For i% = 0 To last_angle% - 1
  For j% = last_angle% To i% + 1 Step -1
    Call combine_two_angle_with_para(A(i%), A(j%), 0, 0, p(i%), p(j%), v$, "", 0, 0, 0, _
         temp_record.record_data)
  Next j%
 Next i%
 i% = last_angle%
 Call remove_record_for_zero_para(p(), A(), last_angle%)
 If i% = last_angle% Or last_angle% = -1 Then
  Exit Function
 Else
  combine_six_angle0 = combine_six_angle0(A(0), A(1), A(2), A(3), A(4), A(5), _
      p(0), p(1), p(2), p(3), p(4), p(5), v$, last_angle%, temp_record.record_data)
 End If
End If
End Function
Public Function combine_dpoint_pair_with_eline(ByVal dp%, _
                       ByVal start%, ByVal no_reduce As Byte) As Byte '10.10
Dim i%, k%, l%, no%
Dim n(3) As Integer
Dim m(3) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim last_tn%
Dim el As eline_data0_type
If Ddpoint_pair(dp%).record_.no_reduce > 4 Then
 Exit Function
End If
For k% = 0 To 3
n(0) = k%
n(1) = (k% + 1) Mod 4
n(2) = (k% + 2) Mod 4
n(3) = (k% + 3) Mod 4
For l% = 0 To 1
m(0) = l%
m(1) = (l% + 1) Mod 2
el.poi(2 * m(0)) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n(0))
el.poi(2 * m(0) + 1) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n(0) + 1)
el.poi(2 * m(1)) = -1
Call search_for_eline(el, m(0), n_(0), 1)  '5.7
el.poi(2 * m(1)) = 30000
Call search_for_eline(el, m(0), n_(1), 1)
last_tn% = 0
For i% = n_(0) + 1 To n_(1)
no% = Deline(i%).data(0).record.data1.index.i(m(0))  '5.7
If no% > start% Then
 'If is_two_record_related(eline_, no%, Deline(no%).data(0).record, _
      dpoint_pair_, dp%, Ddpoint_pair(dp%).data(0).record) = False And _
       Deline(no%).record_.no_reduce < 255 Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
'End If
Next i%
For i% = 1 To last_tn%
no% = tn(i%)
combine_dpoint_pair_with_eline = _
    combine_relation_with_dpoint_pair_(eline_, no%, dp%, m(0), n(0), no_reduce)
 If combine_dpoint_pair_with_eline > 1 Then
  Exit Function
 End If
Next i%
Next l%
Next k%
End Function
Public Function combine_dpoint_pair_with_item(dp%, no_reduce As Byte) As Byte '10.10
Dim i%, j%, k%, no%, last_tn%
Dim n(3) As Integer
Dim m(2) As Integer
Dim ite As item0_data_type
Dim tn() As Integer
Dim n_(1) As Integer
For i% = 0 To 3
n(0) = i%
 n(1) = (i% + 1) Mod 4
For j% = 0 To 2
 m(0) = j%
  m(1) = (j% + 1) Mod 3
   m(2) = (j% + 1) Mod 3
ite.poi(2 * m(0)) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n(0))
ite.poi(2 * m(0) + 1) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n(0) + 1)
ite.poi(2 * m(1)) = -1
Call search_for_item0(ite, j%, n_(0), 1)
ite.poi(2 * m(1)) = 30000
Call search_for_item0(ite, j%, n_(1), 1)  '5.7
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
no% = item0(k%).data(0).index(m(0))
If no% > 0 Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
combine_dpoint_pair_with_item = _
  combine_item_with_point_pair_(no%, dp%, m(0), n(0), no_reduce)
If combine_dpoint_pair_with_item > 1 Then
 Exit Function
End If
Next k%
Next j%
Next i%
End Function
Public Function combine_dpoint_pair_with_line_value(ByVal dp%, _
                           ByVal start%, ByVal no_reduce As Byte) As Byte '10.10
Dim i%, no%, tn%
Dim n(3) As Integer
If Ddpoint_pair(dp%).record_.no_reduce > 4 Then
 Exit Function
End If
For i% = 0 To 3
n(0) = i%
n(1) = (i% + 1) Mod 4
n(2) = (i% + 2) Mod 4
n(3) = (i% + 3) Mod 4
If is_line_value(Ddpoint_pair(dp%).data(0).data0.poi(2 * n(0)), _
    Ddpoint_pair(dp%).data(0).data0.poi(2 * n(0) + 1), _
     Ddpoint_pair(dp%).data(0).data0.n(2 * n(1)), _
      Ddpoint_pair(dp%).data(0).data0.n(2 * n(0) + 1), _
       Ddpoint_pair(dp%).data(0).data0.line_no(n(0)), _
        "", tn%, -1000, 0, 0, 0, line_value_data0) = 1 Then
If tn% > start% Then
  If line_value(tn%).record_.no_reduce < 255 Then
   combine_dpoint_pair_with_line_value = _
    combine_relation_with_dpoint_pair_(line_value_, tn%, _
     dp%, 0, n(0), no_reduce)
If combine_dpoint_pair_with_line_value > 0 Then
   Call set_level_(Ddpoint_pair(dp%).record_.no_reduce, 4)
End If
  If combine_dpoint_pair_with_line_value > 1 Then
   Exit Function
  End If
End If
End If
End If
Next i%
End Function
Public Function combine_dpoint_pair_with_mid_point(ByVal dp%, _
              ByVal start%, ByVal no_reduce As Byte) As Byte '10.10
Dim i%, k%, l%, no%
Dim n(3) As Integer
Dim m(5) As Integer
Dim v(1) As String
Dim mdp As mid_point_data0_type
If Ddpoint_pair(dp%).record_.no_reduce > 4 Then
 Exit Function
End If
For k% = 0 To 3
n(0) = k%
n(1) = (k% + 1) Mod 4
n(2) = (k% + 2) Mod 4
n(3) = (k + 3) Mod 4
For l% = 0 To 2
If l% = 0 Then
m(0) = 0
m(1) = 1
m(2) = 1
m(3) = 2
m(4) = 0
m(5) = 2
ElseIf l% = 1 Then
m(0) = 1
m(1) = 2
m(2) = 0
m(3) = 2
m(4) = 0
m(5) = 1
Else
m(0) = 0
m(1) = 2
m(2) = 0
m(3) = 1
m(4) = 1
m(5) = 2
End If
mdp.poi(m(0)) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n(0))
mdp.poi(m(1)) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n(0) + 1)
If search_for_mid_point(mdp, l%, no%, 2) Then   '5.7原l%+3
If no% > start% And Dmid_point(no%).record_.no_reduce < 4 Then
'If is_two_record_related(midpoint_, no%, Dmid_point(no%).data(0).record, _
     dpoint_pair_, dp%, Ddpoint_pair(dp%).data(0).record) = False Then
  combine_dpoint_pair_with_mid_point = _
   combine_relation_with_dpoint_pair_(midpoint_, no%, dp%, l%, n(0), no_reduce)
 If combine_dpoint_pair_with_mid_point > 1 Then
   Exit Function
 End If
 End If
'End If
End If
Next l%
Next k%
End Function

Public Function combine_line_value_with_dpoint_pair(ByVal lv%, _
            ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, no%, tn%
Dim it(1) As Integer
Dim m(3) As Integer
Dim n_(1) As Integer
Dim tn0() As Integer
Dim tn1() As Integer
Dim last_tn0%, last_tn1%
Dim para(1) As String
Dim ddp As point_pair_data0_type
Dim temp_record As total_record_type
Dim re As total_record_type
Call add_conditions_to_record(line_value_, lv%, 0, 0, re.record_data.data0.condition_data)
re.record_data.data0.theorem_no = 1
For i% = 0 To 3
m(0) = i%
m(1) = (i% + 1) Mod 4
m(2) = (i% + 2) Mod 4
m(3) = (i% + 3) Mod 4
ddp.poi(2 * m(0)) = line_value(lv%).data(0).data0.poi(0)
ddp.poi(2 * m(0) + 1) = line_value(lv%).data(0).data0.poi(1)
ddp.poi(2 * m(1)) = -1
Call search_for_point_pair(ddp, m(0), n_(0), 1)
ddp.poi(2 * m(1)) = 30000
Call search_for_point_pair(ddp, m(0), n_(1), 1)  '5.7
If m(0) = 1 Then
m(1) = 0
m(2) = 3
m(3) = 2
ElseIf m(0) = 3 Then
m(1) = 2
m(2) = 1
m(3) = 0
End If
last_tn0% = 0
last_tn1% = 0
For j% = n_(0) + 1 To n_(1)
no% = Ddpoint_pair(j%).data(0).record.data1.index.i(m(0))
If no% > start% And Ddpoint_pair(no%).record_.no_reduce < 4 Then
 If Ddpoint_pair(no%).data(0).data0.poi(2 * m(3)) = line_value(lv%).data(0).data0.poi(0) And _
      Ddpoint_pair(no%).data(0).data0.poi(2 * m(3) + 1) = line_value(lv%).data(0).data0.poi(1) Then
 If m(0) < 2 Then
 last_tn1% = last_tn1% + 1
 ReDim Preserve tn1(last_tn1%) As Integer
 tn1(last_tn1%) = no%
 End If
 Else
 last_tn0% = last_tn0% + 1
 ReDim Preserve tn0(last_tn0%) As Integer
 tn0(last_tn0%) = no%
 End If
End If
Next j%
For j% = 1 To last_tn0%
no% = tn0(j%)
temp_record = re
Call add_conditions_to_record(dpoint_pair_, no%, 0, 0, temp_record.record_data.data0.condition_data)
combine_line_value_with_dpoint_pair = set_general_string_from_relation( _
 0, 0, Ddpoint_pair(no%).data(0).data0.poi(2 * m(1)), _
 Ddpoint_pair(no%).data(0).data0.poi(2 * m(1) + 1), _
  Ddpoint_pair(no%).data(0).data0.poi(2 * m(2)), _
   Ddpoint_pair(no%).data(0).data0.poi(2 * m(2) + 1), _
    Ddpoint_pair(no%).data(0).data0.poi(2 * m(3)), _
     Ddpoint_pair(no%).data(0).data0.poi(2 * m(3) + 1), _
 0, 0, Ddpoint_pair(no%).data(0).data0.n(2 * m(1)), _
 Ddpoint_pair(no%).data(0).data0.n(2 * m(1) + 1), _
  Ddpoint_pair(no%).data(0).data0.n(2 * m(2)), _
   Ddpoint_pair(no%).data(0).data0.n(2 * m(2) + 1), _
    Ddpoint_pair(no%).data(0).data0.n(2 * m(3)), _
     Ddpoint_pair(no%).data(0).data0.n(2 * m(3) + 1), _
   0, Ddpoint_pair(no%).data(0).data0.line_no(m(1)), _
    Ddpoint_pair(no%).data(0).data0.line_no(m(2)), _
     Ddpoint_pair(no%).data(0).data0.line_no(m(3)), _
      line_value(lv%).data(0).data0.value, "1", temp_record, no_reduce)
  Call set_level_(Ddpoint_pair(no%).record_.no_reduce, 4)
If combine_line_value_with_dpoint_pair > 1 Then
Exit Function
End If
Next j%
For j% = 1 To last_tn1%
no% = tn1(j%)
temp_record = re
Call add_conditions_to_record(dpoint_pair_, no%, 0, 0, temp_record.record_data.data0.condition_data)
combine_line_value_with_dpoint_pair = set_item0(Ddpoint_pair(no%).data(0).data0.poi(2 * m(1)), _
 Ddpoint_pair(no%).data(0).data0.poi(2 * m(1) + 1), _
   Ddpoint_pair(no%).data(0).data0.poi(2 * m(2)), _
    Ddpoint_pair(no%).data(0).data0.poi(2 * m(2) + 1), _
    "*", Ddpoint_pair(no%).data(0).data0.n(2 * m(1)), _
     Ddpoint_pair(no%).data(0).data0.n(2 * m(1) + 1), _
      Ddpoint_pair(no%).data(0).data0.n(2 * m(2)), _
       Ddpoint_pair(no%).data(0).data0.n(2 * m(2) + 1), _
        Ddpoint_pair(no%).data(0).data0.line_no(m(1)), _
         Ddpoint_pair(no%).data(0).data0.line_no(m(2)), _
          "1", "1", "1", "", para(0), 0, record_data0.data0.condition_data, _
             0, it(0), no_reduce, 0, condition_data0, False)
          If combine_line_value_with_dpoint_pair > 1 Then
             Exit Function
          End If
If it(0) > 0 Then
If is_line_value(Ddpoint_pair(no%).data(0).data0.poi(2 * m(3)), _
    Ddpoint_pair(no%).data(0).data0.poi(2 * m(3) + 1), _
     Ddpoint_pair(no%).data(0).data0.n(2 * m(3)), _
      Ddpoint_pair(no%).data(0).data0.n(2 * m(3) + 1), _
       Ddpoint_pair(no%).data(0).data0.line_no(m(3)), _
        "", tn%, -1000, 0, 0, 0, line_value_data0) = 1 Then
' Call add_conditions_to_record(line_value_, tn%, 0, 0, temp_record.record_data)
'combine_line_value_with_dpoint_pair = set_general_string( _
 it(0), 0, 0, 0, "1", "0", "0", "0", time_string( _
   line_value(lv%).data(0).data0.value, line_value(tn%).data(0).data0.value), _
  0, 1, temp_record, 0, no_reduce)
   Call set_level_(Ddpoint_pair(no%).record_.no_reduce, 4)
'If combine_line_value_with_dpoint_pair > 1 Then
'Exit Function
'End If
Else
'Call set_item0(Ddpoint_pair(no%).data(0).data0.poi(2 * m(3)), _
  Ddpoint_pair(no%).data(0).data0.poi(2 * m(3) + 1), _
   0, 0, "~", Ddpoint_pair(no%).data(0).data0.n(2 * m(3)), _
     Ddpoint_pair(no%).data(0).data0.n(2 * m(3) + 1), 0, 0, _
        Ddpoint_pair(no%).data(0).data0.line_no(m(3)), 0, _
         "", record_00, "", record_data0, 0, it(1))
'combine_line_value_with_dpoint_pair = set_general_string( _
' it(0), it(1), 0, 0, "1", "-1", "0", "0", "0", _
'  0, 1, temp_record, 0, no_reduce)
'   Call set_level_(Ddpoint_pair(no%).record_.no_reduce, 4)
'If combine_line_value_with_dpoint_pair > 1 Then
'Exit Function
'End If
End If
End If
Next j%
Next i%
End Function
Public Function combine_relation_with_item_(ByVal ty As Byte, _
      ByVal re%, ByVal it%, k%, j%, no_reduce As Byte) As Byte         'l%替换的线段号
Dim tn%, t_n%
Dim m(2) As Integer
Dim p(5) As Integer
Dim n(5) As Integer
Dim l(3) As Integer
Dim temp_record As record_type0
Dim temp_record0_ As record_data_type
Dim temp_record_ As total_record_type
Dim v(1) As String
Dim tv As String
temp_record_.record_data.data0.condition_data.condition_no = 1
 temp_record_.record_data.data0.condition_data.condition(1).ty = ty
  temp_record_.record_data.data0.condition_data.condition(1).no = re%
   temp_record_.record_data.data0.theorem_no = 1
temp_record0_.data0.condition_data.condition_no = 1
 temp_record0_.data0.condition_data.condition(1).ty = ty
  temp_record0_.data0.condition_data.condition(1).no = re%
temp_record.condition_data = temp_record0_.data0.condition_data
If item0(it%).data(0).no_reduce = False Then
Call read_point_and_ratio_from_relation(ty, re%, k%, p(), n(), l(), v(0), v(1)) '读出比值
m(0) = j%
   If item0(it%).data(0).poi(4) > 0 And item0(it%).data(0).poi(5) > 0 Then
    m(1) = (m(0) + 1) Mod 3
     m(2) = (m(0) + 2) Mod 3
   Else
    m(1) = (m(0) + 1) Mod 2
     m(2) = 2
   End If
'Else '<>line_value
temp_record.condition_data = temp_record0_.data0.condition_data
 If item0(it%).data(0).sig = "~" Then
  If j% = 0 Then
      temp_record.condition_data.condition(3).no = (k% + 1) Mod 3
       temp_record.para(0) = v(0)
    combine_relation_with_item_ = set_item0(p(2), p(3), _
       0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", v(0), "", "1", 0, _
        temp_record0_.data0.condition_data, it%, t_n%, no_reduce, 0, condition_data0, False)
    If combine_relation_with_item_ > 1 Then
     Exit Function
    End If
    If v(1) <> "" > 0 Then
       temp_record.condition_data.condition(3).no = (k% + 2) Mod 3
     combine_relation_with_item_ = set_item0(p(4), p(5), _
       0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", v(1), "", "1", 0, _
         temp_record0_.data0.condition_data, it%, t_n%, no_reduce, 0, condition_data0, False)
    If combine_relation_with_item_ > 1 Then
     Exit Function
    End If
    End If
   End If
 ElseIf item0(it%).data(0).sig = "*" Then
   If j% = 2 Then
      If item0(it%).data(0).value <> "" Then
      
      End If
   ElseIf j% < 2 Then
    m(1) = (j% + 1) Mod 2
     If item0(it%).data(0).poi(2 * m(1)) = p(0) And _
          item0(it%).data(0).poi(2 * m(1) + 1) = p(1) Then
      If v(0) <> "" Then
       tv = time_string(v(0), v(0), True, False)
         temp_record.condition_data.condition(3).no = (k% + 1) Mod 3
        If item0(it%).data(0).value = "" Then
          combine_relation_with_item_ = set_item0(p(2), p(3), _
            p(2), p(3), "*", n(2), n(3), n(2), n(3), l(1), l(1), _
             "1", "1", tv, "", "1", 0, temp_record0_.data0.condition_data, _
               it%, t_n%, no_reduce, 0, condition_data0, False)
          If combine_relation_with_item_ > 1 Then
           Exit Function
          End If
        Else
         tv = divide_string(item0(it%).data(0).value, tv, True, False)
          tv = sqr_string(tv, True, False)
           temp_record_.record_data.data0.condition_data = temp_record.condition_data
            Call add_record_to_record(item0(it%).data(0).record_for_value.data0.condition_data, _
                temp_record_.record_data.data0.condition_data)
           combine_relation_with_item_ = set_line_value(p(2), p(3), tv, _
            n(2), n(3), l(1), temp_record_, 0, 0, False)
          If combine_relation_with_item_ > 1 Then
           Exit Function
          End If
        End If
     End If 'v(0)<>""
      '*********************
     If v(1) <> "" Then
      tv = time_string(v(1), v(1), True, False)
       temp_record.condition_data.condition(3).no = (k% + 2) Mod 3
      If item0(it%).data(0).value = "" Then
        combine_relation_with_item_ = set_item0(p(4), p(5), _
         p(4), p(5), "*", n(4), n(5), n(4), n(5), l(2), l(2), _
           "1", "1", tv, "", "1", 0, temp_record0_.data0.condition_data, it%, _
              t_n%, no_reduce, 0, condition_data0, False)
         If combine_relation_with_item_ > 1 Then
          Exit Function
         End If
      Else
       temp_record_.record_data.data0.condition_data = temp_record.condition_data
        tv = divide_string(item0(it%).data(0).value, tv, False, False)
         tv = sqr_string(tv, True, False)
         Call add_record_to_record(item0(it%).data(0).record_for_value.data0.condition_data, _
                temp_record_.record_data.data0.condition_data)
         combine_relation_with_item_ = set_line_value(p(4), p(5), tv, _
           n(4), n(5), l(2), temp_record_, 0, 0, False)
         If combine_relation_with_item_ > 1 Then
          Exit Function
         End If
      End If
     End If 'v(1)<>""
    ElseIf item0(it%).data(0).poi(2 * m(1)) = p(2) And _
         item0(it%).data(0).poi(2 * m(1) + 1) = p(3) Then
     If v(0) <> "" Then
      tv = divide_string("1", v(0), True, False)
      temp_record.condition_data.condition(3).no = k%
      If item0(it%).data(0).value = "" Then
       combine_relation_with_item_ = set_item0(p(0), p(1), _
        p(0), p(1), "*", n(0), n(1), n(0), n(1), l(0), l(0), _
         "1", "1", tv, "", "1", 0, temp_record0_.data0.condition_data, it%, _
           t_n%, no_reduce, 0, condition_data0, False)
       If combine_relation_with_item_ > 1 Then
        Exit Function
       End If
      Else
       temp_record_.record_data.data0.condition_data = temp_record.condition_data
        Call add_record_to_record(item0(it%).data(0).record_for_value.data0.condition_data, _
                    temp_record_.record_data.data0.condition_data)
        tv = divide_string(item0(it%).data(0).value, tv, False, False)
         tv = sqr_string(tv, True, False)
        combine_relation_with_item_ = set_line_value(p(0), p(1), _
         tv, n(0), n(1), l(0), temp_record_, 0, 0, False)
         If combine_relation_with_item_ > 1 Then
          Exit Function
         End If
      End If
     End If
     If v(0) <> "" Then
      temp_record.condition_data.condition(3).no = (k% + 1) Mod 3
        If item0(it%).data(0).value = "" Then
         combine_relation_with_item_ = set_item0(p(2), p(3), _
          p(2), p(3), "*", n(2), n(3), n(2), n(3), l(1), l(1), _
           "1", "1", v(0), "", "1", 0, temp_record0_.data0.condition_data, _
              it%, t_n%, no_reduce, 0, condition_data0, False)
            If combine_relation_with_item_ > 1 Then
             Exit Function
            End If
        Else
         temp_record_.record_data.data0.condition_data = temp_record.condition_data
          Call add_record_to_record(item0(it%).data(0).record_for_value.data0.condition_data, _
                                                     temp_record_.record_data.data0.condition_data)
           tv = divide_string(item0(it%).data(0).value, v(0), False, False)
            tv = sqr_string(tv, True, False)
             combine_relation_with_item_ = set_line_value(p(2), p(3), _
              tv, n(2), n(3), l(1), temp_record_, 0, 0, False)
           If combine_relation_with_item_ > 1 Then
             Exit Function
           End If
        End If
     End If
     If v(0) <> "" And v(1) <> "" Then
      tv = divide_string(time_string(v(1), v(1), False, False), v(0), True, False)
       temp_record.condition_data.condition(3).no = (k% + 2) Mod 3
      If item0(it%).data(0).value = "" Then
       combine_relation_with_item_ = set_item0(p(4), p(5), _
        p(4), p(5), "*", n(4), n(5), n(4), n(5), l(2), l(2), _
         "1", "1", tv, "", "1", 0, temp_record0_.data0.condition_data, _
           it%, t_n%, no_reduce, 0, condition_data0, False)
        If combine_relation_with_item_ > 1 Then
         Exit Function
        End If
      Else
       tv = divide_string(item0(it%).data(0).value, tv, False, False)
        tv = sqr_string(tv, True, False)
         temp_record_.record_data.data0.condition_data = temp_record.condition_data
          Call add_record_to_record(item0(it%).data(0).record_for_value.data0.condition_data, _
                temp_record_.record_data.data0.condition_data)
         combine_relation_with_item_ = set_line_value(p(4), p(5), _
          tv, n(4), n(5), l(2), temp_record_, 0, 0, False)
           If combine_relation_with_item_ > 1 Then
            Exit Function
           End If
      End If
    End If
   ElseIf item0(it%).data(0).poi(2 * m(1)) = p(4) And _
         item0(it%).data(0).poi(2 * m(1) + 1) = p(5) Then
     If v(1) <> "" Then
      tv = divide_string("1", v(1), True, False)
       temp_record.condition_data.condition(3).no = k%
        If item0(it%).data(0).value = "" Then
         combine_relation_with_item_ = set_item0(p(0), p(1), _
          p(0), p(1), "*", n(0), n(1), n(0), n(1), l(0), l(0), _
           "1", "1", tv, "", "1", 0, temp_record0_.data0.condition_data, _
             it%, t_n%, no_reduce, 0, condition_data0, False)
           If combine_relation_with_item_ > 1 Then
             Exit Function
           End If
        Else
         tv = divide_string(item0(it%).data(0).value, tv, False, False)
          tv = sqr_string(tv, True, False)
         temp_record_.record_data.data0.condition_data = temp_record.condition_data
          Call add_record_to_record(item0(it%).data(0).record_for_value.data0.condition_data, _
                                                temp_record_.record_data.data0.condition_data)
           combine_relation_with_item_ = set_line_value(p(0), p(1), _
            tv, n(0), n(1), l(0), temp_record_, 0, 0, False)
          If combine_relation_with_item_ > 1 Then
           Exit Function
          End If
        End If
     End If
     If v(0) <> "" And v(1) <> "" Then
      tv = divide_string(time_string(v(0), v(0), False, False), v(1), True, False)
       temp_record.condition_data.condition(3).no = (k% + 1) Mod 3
        If item0(it%).data(0).value = "" Then
         combine_relation_with_item_ = set_item0(p(2), p(3), _
          p(2), p(3), "*", n(2), n(3), n(2), n(3), l(1), l(1), _
           "1", "1", tv, "", "1", 0, temp_record0_.data0.condition_data, it%, _
             t_n%, no_reduce, 0, condition_data0, False)
            If combine_relation_with_item_ > 1 Then
              Exit Function
            End If
         Else
          temp_record_.record_data.data0.condition_data = temp_record.condition_data
           Call add_record_to_record(item0(it%).data(0).record_for_value.data0.condition_data, _
                                     temp_record_.record_data.data0.condition_data)
            tv = divide_string(item0(it%).data(0).value, tv, False, False)
             tv = sqr_string(tv, True, False)
           combine_relation_with_item_ = set_line_value(p(2), p(3), tv, _
            n(2), n(3), l(1), temp_record_, 0, 0, False)
             If combine_relation_with_item_ > 1 Then
              Exit Function
             End If
         End If
     End If
     If v(1) <> "" Then
     temp_record.condition_data.condition(3).no = (k% + 2) Mod 3
      If item0(it%).data(0).value = "" Then
       combine_relation_with_item_ = set_item0(p(4), p(5), _
        p(4), p(5), "*", n(4), n(5), n(4), n(5), l(2), l(2), _
         "1", "1", v(1), "", "1", 0, temp_record0_.data0.condition_data, it%, _
            t_n%, no_reduce, 0, condition_data0, False)
         If combine_relation_with_item_ > 1 Then
          Exit Function
         End If
      Else
       tv = divide_string(item0(it%).data(0).value, v(1), False, False)
        tv = sqr_string(tv, True, False)
         temp_record_.record_data.data0.condition_data = temp_record.condition_data
          Call add_record_to_record(item0(it%).data(0).record_for_value.data0.condition_data, _
                                          temp_record_.record_data.data0.condition_data)
          combine_relation_with_item_ = set_line_value(p(4), p(5), tv, _
            n(4), n(5), l(2), temp_record_, 0, 0, False)
            If combine_relation_with_item_ > 1 Then
             Exit Function
            End If
      End If
 '  temp_record.record_data.data0.condition_data.condition(3).no = 4
 '    combine_relation_with_item_ = add_new_item_to_item(t_n%, 0, _
       v(1), "0", it%, temp_record.record_data)
 '    If combine_relation_with_item_ > 1 Then
 '     Exit Function
 '    End If
     End If
   Else
     If v(0) <> "" Then
      temp_record.condition_data.condition(3).no = (k% + 1) Mod 3
       combine_relation_with_item_ = set_item0(p(2), p(3), _
        item0(it%).data(0).poi(2 * m(1)), item0(it%).data(0).poi(2 * m(1) + 1), _
         "*", n(2), n(3), item0(it%).data(0).n(2 * m(1)), item0(it%).data(0).n(2 * m(1) + 1), _
          l(1), item0(it%).data(0).line_no(m(1)), "1", "1", v(0), "", "1", 0, _
            temp_record0_.data0.condition_data, it%, t_n%, no_reduce, 0, condition_data0, False)
      If combine_relation_with_item_ > 1 Then
       Exit Function
      End If
     End If
     If v(1) <> "" Then
      temp_record.condition_data.condition(3).no = (k% + 2) Mod 3
       combine_relation_with_item_ = set_item0(p(4), p(5), _
        item0(it%).data(0).poi(2 * m(1)), item0(it%).data(0).poi(2 * m(1) + 1), _
         "*", n(4), n(5), item0(it%).data(0).n(2 * m(1)), item0(it%).data(0).n(2 * m(1) + 1), _
          l(2), item0(it%).data(0).line_no(m(1)), "1", "1", v(1), "", "1", 0, _
             temp_record0_.data0.condition_data, it%, t_n%, no_reduce, 0, condition_data0, False)
      If combine_relation_with_item_ > 1 Then
       Exit Function
      End If
    End If
    End If
    End If
ElseIf item0(it%).data(0).sig = "/" Then
  If v(0) <> "" Then
   If item0(it%).data(0).poi(2 * m(1)) = p(2) And item0(it%).data(0).poi(2 * m(1) + 1) = p(3) Then
      temp_record.condition_data.condition(3).no = 4
     combine_relation_with_item_ = set_item0_value(it%, m(0), m(1), _
        v(0), "1", "", "", 0, temp_record.condition_data)
        If combine_relation_with_item_ > 1 Then
         Exit Function
        End If
   ElseIf item0(it%).data(0).poi(2 * m(2)) = p(2) And item0(it%).data(0).poi(2 * m(2) + 1) = p(3) Then
   temp_record.condition_data.condition(3).no = 4
     combine_relation_with_item_ = set_item0_value(it%, m(0), m(2), _
        v(0), "1", "", "", 0, temp_record.condition_data)
        If combine_relation_with_item_ > 1 Then
         Exit Function
        End If
   Else
     If m(0) = 0 Then
    temp_record.condition_data.condition(3).no = (k% + 1) Mod 3
       combine_relation_with_item_ = set_item0(p(2), p(3), item0(it%).data(0).poi(2), _
          item0(it%).data(0).poi(3), "/", n(2), n(3), item0(it%).data(0).n(2), item0(it%).data(0).n(3), _
           l(1), item0(it%).data(0).line_no(1), "1", "1", v(0), "", "1", 0, _
             temp_record0_.data0.condition_data, it%, t_n%, no_reduce, 0, condition_data0, False)
    If combine_relation_with_item_ > 1 Then
     Exit Function
    End If
     ElseIf m(0) = 1 Then
      tv = divide_string("1", v(0), True, False)
    temp_record.condition_data.condition(3).no = (k% + 1) Mod 3
       combine_relation_with_item_ = set_item0(item0(it%).data(0).poi(0), _
          item0(it%).data(0).poi(1), p(2), p(3), "/", item0(it%).data(0).n(0), _
           item0(it%).data(0).n(1), n(2), n(3), item0(it%).data(0).line_no(0), l(1), _
            "1", "1", tv, "", "1", 0, temp_record0_.data0.condition_data, _
              it%, t_n%, no_reduce, 0, condition_data0, False)
    If combine_relation_with_item_ > 1 Then
     Exit Function
    End If
     End If
   End If
  ElseIf v(1) <> "" Then
   If item0(it%).data(0).poi(2 * m(1)) = p(4) And item0(it%).data(0).poi(2 * m(1) + 1) = p(5) Then
    temp_record.condition_data.condition(3).no = 4
      combine_relation_with_item_ = set_item0_value(it%, m(0), m(1), _
         v(1), "1", "", "", 0, temp_record.condition_data)
        If combine_relation_with_item_ > 1 Then
         Exit Function
        End If
           item0(it%).data(0).no_reduce = True
             Call combine_item_with_general_string_(it%, -2)
  ElseIf item0(it%).data(0).poi(2 * m(2)) = p(4) And item0(it%).data(0).poi(2 * m(2) + 1) = p(5) Then
    temp_record.condition_data.condition(3).no = 4
     combine_relation_with_item_ = set_item0_value(it%, m(0), m(2), _
        v(1), "1", "", "", 0, temp_record.condition_data)
     If combine_relation_with_item_ > 1 Then
      Exit Function
     End If
             item0(it%).data(0).no_reduce = True
             Call combine_item_with_general_string_(it%, -2)
   Else
     If m(0) = 0 Then
   temp_record.condition_data.condition(3).no = (k% + 1) Mod 3
        combine_relation_with_item_ = set_item0(p(2), p(3), item0(it%).data(0).poi(2), _
          item0(it%).data(0).poi(3), "/", n(2), n(3), item0(it%).data(0).n(2), item0(it%).data(0).n(3), _
           l(1), item0(it%).data(0).line_no(1), "1", "1", v(1), "", "1", 0, _
             temp_record0_.data0.condition_data, it%, t_n%, no_reduce, 0, condition_data0, False)
    If combine_relation_with_item_ > 1 Then
     Exit Function
    End If
     ElseIf m(0) = 1 Then
     tv = divide_string("1", v(1), True, False)
     temp_record.condition_data.condition(3).no = (k% + 1) Mod 3
      combine_relation_with_item_ = set_item0(item0(it%).data(0).poi(0), _
          item0(it%).data(0).poi(1), p(2), p(3), "/", item0(it%).data(0).n(0), _
           item0(it%).data(0).n(1), n(2), n(3), item0(it%).data(0).line_no(0), l(1), "1", "1", tv, _
             "", "1", 0, temp_record0_.data0.condition_data, it%, t_n%, no_reduce, 0, condition_data0, False)
    If combine_relation_with_item_ > 1 Then
     Exit Function
    End If
    End If
   End If
  End If
  End If
 End If
 End Function
Public Function combine_item_with_line_value(ByVal it%, no_reduce As Byte) As Byte
Dim i%, tn%, tn1%, tn2%, it_%
Dim pA$
Dim temp_record As record_data_type
Dim temp_record1 As condition_data_type
Dim t_it%
Dim lv_data As line_value_data0_type
If is_line_value(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), item0(it%).data(0).n(0), _
     item0(it%).data(0).n(1), item0(it%).data(0).line_no(0), "", _
      tn%, -1000, 0, 0, 0, line_value_data0) = 1 Then
        temp_record.data0.condition_data.condition_no = 1
         temp_record.data0.condition_data.condition(1).no = tn%
          temp_record.data0.condition_data.condition(1).ty = line_value_
  If item0(it%).data(0).poi(2) = 0 And item0(it%).data(0).poi(3) = 0 Then
   combine_item_with_line_value = _
      set_item0_value(it%, 0, 1, line_value(tn%).data(0).data0.value, "0", "", "", 1, temp_record.data0.condition_data)
    If combine_item_with_line_value > 1 Then
     Exit Function
    End If
  Else
  If is_line_value(item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), item0(it%).data(0).n(2), _
    item0(it%).data(0).n(3), item0(it%).data(0).line_no(1), "", _
      tn1%, -1000, 0, 0, 0, line_value_data0) = 1 Then
       Call add_conditions_to_record(line_value_, tn1%, 0, 0, temp_record.data0.condition_data)
   combine_item_with_line_value = set_item0_value(it%, 0, 1, line_value(tn%).data(0).data0.value, _
           line_value(tn1%).data(0).data0.value, "", "", 1, temp_record.data0.condition_data)
    If combine_item_with_line_value > 1 Then
     Exit Function
    End If
  Else
   combine_item_with_line_value = combine_line_value_with_item_(tn%, it%, 0, no_reduce)
    If combine_item_with_line_value > 1 Then
     Exit Function
    End If
  End If
  End If
ElseIf is_line_value(item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), item0(it%).data(0).n(2), _
    item0(it%).data(0).n(3), item0(it%).data(0).line_no(1), "", _
      tn%, -1000, 0, 0, 0, line_value_data0) = 1 Then
  combine_item_with_line_value = combine_line_value_with_item_(tn%, it%, 1, no_reduce)
ElseIf is_line_value(item0(it%).data(0).poi(4), item0(it%).data(0).poi(5), item0(it%).data(0).n(4), _
      item0(it%).data(0).n(5), item0(it%).data(0).line_no(2), "", _
      tn%, -1000, 0, 0, 0, line_value_data0) = 1 Then
  combine_item_with_line_value = combine_line_value_with_item_(tn%, it%, 2, no_reduce)
End If
If item0(it%).data(0).sig = "~" And item0(it%).data(0).value = "" Then
   For i% = item0(it%).data(0).n(0) + 1 To item0(it%).data(0).n(1) - 1
    If is_line_value(item0(it%).data(0).poi(0), m_lin(item0(it%).data(0).line_no(0)).data(0).data0.in_point(i%), _
       0, 0, 0, "", tn%, -1000, 0, 0, 0, lv_data) = 1 Then
        temp_record1.condition_no = 1
         temp_record1.condition(1).no = tn%
          temp_record1.condition(1).ty = line_value_
        Call set_item0(m_lin(item0(it%).data(0).line_no(0)).data(0).data0.in_point(i%), _
          item0(it%).data(0).poi(1), 0, 0, "~", 0, 0, _
           0, 0, 0, 0, "1", "1", "1", "", "1", 0, temp_record.data0.condition_data, _
            0, it_%, 0, 0, condition_data0, False)
        combine_item_with_line_value = add_new_item_to_item(it_%, 0, "1", _
             line_value(tn%).data(0).data0.value_, it%, temp_record1)
        If combine_item_with_line_value > 1 Then
           Exit Function
        End If
    ElseIf is_line_value(m_lin(item0(it%).data(0).line_no(0)).data(0).data0.in_point(i%), _
            item0(it%).data(0).poi(1), 0, 0, 0, _
           "", tn%, -1000, 0, 0, 0, lv_data) = 1 Then
         temp_record1.condition_no = 1
         temp_record1.condition(1).no = tn%
          temp_record1.condition(1).ty = line_value_
         Call set_item0(item0(it%).data(0).poi(0), _
          m_lin(item0(it%).data(0).line_no(0)).data(0).data0.in_point(i%), 0, 0, "~", 0, 0, _
           0, 0, 0, 0, "1", "1", "1", "", "1", 0, temp_record.data0.condition_data, 0, _
             it_%, 0, 0, condition_data0, False)
        combine_item_with_line_value = add_new_item_to_item(it_%, 0, "1", _
             line_value(tn%).data(0).data0.value_, it%, temp_record1)
        If combine_item_with_line_value > 1 Then
           Exit Function
        End If
  End If
   Next i%
End If
End Function
Public Function combine_item_with_general_string_(ByVal it%, ByVal trans_to_no%) As Byte
Dim i%, j%, k%, n%, no%, it1%, it2%, t_n%
Dim pA(4) As String
Dim ite(4) As Integer
Dim m(3) As Integer
Dim v$
Dim n_(1) As Integer
Dim is_zero1 As Byte
Dim tn() As Integer
Dim last_tn%
Dim ge As general_string_data_type
Dim temp_record As total_record_type
For i% = 0 To 3
m(0) = i%
m(1) = (i% + 1) Mod 4
m(2) = (i% + 2) Mod 4
m(3) = (i% + 3) Mod 4
ge.item(m(0)) = it%
ge.item(m(1)) = -1
Call search_for_general_string(ge, m(0), n_(0), 1)
ge.item(m(1)) = 30000
Call search_for_general_string(ge, m(0), n_(1), 1)
last_tn% = 0
For j% = n_(0) + 1 To n_(1)
 no% = general_string(j%).data(0).record.data1.index.i(m(0))
If no% > 0 Then 'And general_string(no%).record_.no_reduce < 4 Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
Next j%
For j% = 1 To last_tn%
no% = tn(j%)
If trans_to_no% = 0 Then
 For k% = 1 To item0(it%).data(0).record_for_trans.last_trans_to
  combine_item_with_general_string_ = combine_item_with_general_string0(it%, _
          k%, no%, m(0), 0)
    If combine_item_with_general_string_ > 1 Then
     Exit Function
    End If
 Next k%
Else
 If trans_to_no% = -1 Then
  combine_item_with_general_string_ = set_area_of_triangle_from_item0(it%)
 If combine_item_with_general_string_ > 1 Then
  Exit Function
 End If
 End If
combine_item_with_general_string_ = combine_item_with_general_string0(it%, _
             trans_to_no%, no%, m(0), 0)
If combine_item_with_general_string_ > 1 Then
 Exit Function
End If
End If
Next j%
Next i%
End Function
Public Function combine_line_value_with_item(ByVal lv%, _
         ByVal no_reduce As Byte) As Byte
Dim i%, j%, no%, it_%
Dim t_n(1) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim last_tn% ', last_tn01%, last_tn02%
Dim m(2) As Integer
Dim it As item0_data_type
Dim temp_record As condition_data_type
Dim temp_record1 As condition_data_type
For i% = 0 To 2
m(0) = i%
m(1) = (i% + 1) Mod 3
m(2) = (i% + 2) Mod 3
it.poi(2 * m(0)) = line_value(lv%).data(0).data0.poi(0)
it.poi(2 * m(0) + 1) = line_value(lv%).data(0).data0.poi(1)
it.poi(2 * m(1)) = -1
Call search_for_item0(it, m(0), n_(0), 1)  '5.7
it.poi(2 * m(1)) = 30000
Call search_for_item0(it, m(0), n_(1), 1)
last_tn% = 0
For j% = n_(0) + 1 To n_(1)
no% = item0(j%).data(0).index(m(0))
If no% > 0 Then
   last_tn% = last_tn% + 1
    ReDim Preserve tn(last_tn%) As Integer
     tn(last_tn%) = no%
  End If
Next j%
For j% = 1 To last_tn%
 no% = tn(j%)
  combine_line_value_with_item = combine_line_value_with_item_( _
    lv%, no%, i%, no_reduce)
  If combine_line_value_with_item > 1 Then
   Exit Function
  End If
Next j%
Next i%
For i% = last_conditions.last_cond(0).item0_no + 1 To last_conditions.last_cond(1).item0_no
  If item0(i%).data(0).sig = "~" And item0(i%).data(0).value = "" Then
     If item0(i%).data(0).line_no(0) = line_value(lv%).data(0).data0.line_no Then
      If item0(i%).data(0).n(0) = line_value(lv%).data(0).data0.n(0) Then
         If item0(i%).data(0).n(1) > line_value(lv%).data(0).data0.n(1) Then
          temp_record1.condition_no = 1
          temp_record1.condition(1).ty = line_value_
          temp_record1.condition(1).no = lv%
          Call set_item0(line_value(lv%).data(0).data0.poi(1), item0(i%).data(0).poi(1), 0, 0, _
                "~", line_value(lv%).data(0).data0.n(1), item0(i%).data(0).n(1), 0, 0, _
                  line_value(lv%).data(0).data0.line_no, 0, "1", "1", "1", "", "1", 0, temp_record, _
                    0, it_%, 0, 0, condition_data0, False)
         combine_line_value_with_item = add_new_item_to_item(it_%, 0, "1", _
             line_value(lv%).data(0).data0.value_, i%, temp_record1)
           If combine_line_value_with_item > 1 Then
             Exit Function
           End If
         End If
      ElseIf item0(i%).data(0).n(1) = line_value(lv%).data(0).data0.n(1) Then
         If item0(i%).data(0).n(0) < line_value(lv%).data(0).data0.n(0) Then
          temp_record1.condition_no = 1
          temp_record1.condition(1).ty = line_value_
          temp_record1.condition(1).no = lv%
          Call set_item0(line_value(lv%).data(0).data0.poi(0), item0(i%).data(0).poi(0), 0, 0, _
                "~", line_value(lv%).data(0).data0.n(0), item0(i%).data(0).n(0), 0, 0, _
                  line_value(lv%).data(0).data0.line_no, 0, "1", "1", "1", "", "1", 0, temp_record, _
                    0, it_%, 0, 0, condition_data0, False)
         combine_line_value_with_item = add_new_item_to_item(it_%, 0, "1", _
             line_value(lv%).data(0).data0.value_, i%, temp_record1)
           If combine_line_value_with_item > 1 Then
             Exit Function
           End If
         End If
      End If
     End If
  End If
Next i%
End Function

Public Function combine_general_string_with_item(ByVal ge%, _
      ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%
Dim tn() As Integer
Dim it(3) As Integer
  For i% = 0 To 3
   it(i%) = general_string(ge%).data(0).item(i%)
  Next i%
  For i% = 0 To 3
   If it(i%) > 0 Then
     If item0(it(i%)).data(0).value <> "" Then
       combine_general_string_with_item = _
        combine_item_with_general_string0(it(i%), -1, ge%, _
         i%, no_reduce)
       If combine_general_string_with_item > 1 Then
        Exit Function
       End If
     Else
      For k% = 1 To item0(it(i%)).data(0).record_for_trans.last_trans_to
        combine_general_string_with_item = _
          combine_item_with_general_string0(it(i%), _
                k%, ge%, i%, no_reduce)
        If combine_general_string_with_item > 1 Then
         Exit Function
        End If
       Next k%
     End If
   End If
  Next i%
'   combine_general_string_with_item = _
    combine_general_string_with_different_line(ge%, no_reduce)
End Function
Public Function combine_general_string_with_item_value(ByVal ge%, ByVal it%) As Byte
Dim it1%, it2%, tn%
Dim temp_record As total_record_type
Dim c_data As condition_data_type
If it% > 0 Then
   If item0(it%).data(0).sig = "*" And item0(it%).data(0).value <> "" Then
     For ge% = 1 To last_conditions.last_cond(1).general_string_no
         If general_string(ge%).data(0).para(1) <> "0" And general_string(ge%).data(0).para(2) = "0" Then
            If general_string(ge%).data(0).para(0) = "1" And general_string(ge%).data(0).para(1) = "1" Then
                If item0(general_string(ge%).data(0).item(0)).data(0).sig = "*" And _
                      item0(general_string(ge%).data(0).item(1)).data(0).sig = "*" Then
                       If item0(general_string(ge%).data(0).item(0)).data(0).poi(0) = _
                           item0(general_string(ge%).data(0).item(0)).data(0).poi(2) And _
                             item0(general_string(ge%).data(0).item(0)).data(0).poi(1) = _
                               item0(general_string(ge%).data(0).item(0)).data(0).poi(3) And _
                          item0(general_string(ge%).data(0).item(1)).data(0).poi(0) = _
                           item0(general_string(ge%).data(0).item(1)).data(0).poi(2) And _
                             item0(general_string(ge%).data(0).item(1)).data(0).poi(1) = _
                               item0(general_string(ge%).data(0).item(1)).data(0).poi(3) Then
                                If (item0(general_string(ge%).data(0).item(0)).data(0).poi(0) = _
                                     item0(it%).data(0).poi(0) And _
                                        item0(general_string(ge%).data(0).item(0)).data(0).poi(1) = _
                                      item0(it%).data(0).poi(1) And _
                                         item0(general_string(ge%).data(0).item(1)).data(0).poi(0) = _
                                       item0(it%).data(0).poi(2) And _
                                        item0(general_string(ge%).data(0).item(1)).data(0).poi(1) = _
                                         item0(it%).data(0).poi(3)) Or _
                                  (item0(general_string(ge%).data(0).item(0)).data(0).poi(0) = _
                                     item0(it%).data(0).poi(2) And _
                                        item0(general_string(ge%).data(0).item(0)).data(0).poi(1) = _
                                      item0(it%).data(0).poi(3) And _
                                         item0(general_string(ge%).data(0).item(1)).data(0).poi(0) = _
                                       item0(it%).data(0).poi(0) And _
                                        item0(general_string(ge%).data(0).item(1)).data(0).poi(1) = _
                                         item0(it%).data(0).poi(1)) Then
                         combine_general_string_with_item_value = _
                             combine_general_string_with_item_value0(ge%, it%)
                         If combine_general_string_with_item_value > 1 Then
                           Exit Function
                         End If
                         End If
                End If
            End If
         End If
        End If
     Next ge%
   End If
ElseIf ge% > 0 Then
         If general_string(ge%).data(0).para(1) <> "0" And general_string(ge%).data(0).para(2) = "0" Then
            If general_string(ge%).data(0).para(0) = "1" And general_string(ge%).data(0).para(1) = "1" Then
                If item0(general_string(ge%).data(0).item(0)).data(0).sig = "*" And _
                      item0(general_string(ge%).data(0).item(1)).data(0).sig = "*" Then
                       If item0(general_string(ge%).data(0).item(0)).data(0).poi(0) = _
                           item0(general_string(ge%).data(0).item(0)).data(0).poi(2) And _
                             item0(general_string(ge%).data(0).item(0)).data(0).poi(1) = _
                               item0(general_string(ge%).data(0).item(0)).data(0).poi(3) And _
                         item0(general_string(ge%).data(0).item(1)).data(0).poi(0) = _
                           item0(general_string(ge%).data(0).item(1)).data(0).poi(2) And _
                             item0(general_string(ge%).data(0).item(1)).data(0).poi(1) = _
                               item0(general_string(ge%).data(0).item(1)).data(0).poi(3) Then
                  For it% = 1 To last_conditions.last_cond(1).item0_no
                      If item0(it%).data(0).value <> "" Then
                       If item0(it%).data(0).sig = "*" Then
                        If (item0(general_string(ge%).data(0).item(0)).data(0).poi(0) = _
                           item0(it).data(0).poi(0) And _
                             item0(general_string(ge%).data(0).item(0)).data(0).poi(1) = _
                               item0(it%).data(0).poi(1) And _
                                item0(general_string(ge%).data(0).item(1)).data(0).poi(0) = _
                                  item0(it%).data(0).poi(2) And _
                                   item0(general_string(ge%).data(0).item(1)).data(0).poi(1) = _
                                    item0(it).data(0).poi(3)) Or _
                          (item0(general_string(ge%).data(0).item(0)).data(0).poi(0) = _
                           item0(it).data(0).poi(2) And _
                             item0(general_string(ge%).data(0).item(0)).data(0).poi(1) = _
                               item0(it%).data(0).poi(3) And _
                                item0(general_string(ge%).data(0).item(1)).data(0).poi(0) = _
                                  item0(it%).data(0).poi(0) And _
                                   item0(general_string(ge%).data(0).item(1)).data(0).poi(1) = _
                                    item0(it).data(0).poi(1)) Then
                         combine_general_string_with_item_value = _
                             combine_general_string_with_item_value0(ge%, it%)
                         If combine_general_string_with_item_value > 1 Then
                           Exit Function
                         End If
                        End If
                       End If
                      End If
                    Next it%
                  End If
               End If
             End If
          End If

End If
End Function
Public Function combine_general_string_with_item_value0(ByVal ge%, ByVal it%) As Byte
Dim it1%, it2%, tn%
Dim tv(1) As String
Dim s As String
Dim temp_record As total_record_type
Dim c_data As condition_data_type
                         temp_record.record_data.data0.condition_data.condition_no = 0
                         Call add_conditions_to_record(general_string_, ge%, 0, 0, _
                                            temp_record.record_data.data0.condition_data)
                         Call add_record_to_record(item0(it%).data(0).record_for_value.data0.condition_data, _
                                                temp_record.record_data.data0.condition_data)
                         temp_record.record_data.data0.theorem_no = 1
                         s = sqr_string(add_string(general_string(ge%).data(0).value, _
                                      time_string("2", item0(it%).data(0).value, False, False), _
                                        False, False), True, False)
                         Call solut_2order_equation("1", time_string(s, "-1", True, False), _
                             item0(it%).data(0).value, tv(0), tv(1), False)
                         If item0(it%).data(0).poi(1) > 0 And item0(it%).data(0).poi(3) > 0 Then
                         combine_general_string_with_item_value0 = set_two_line_value( _
                            item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), item0(it%).data(0).poi(2), _
                              item0(it%).data(0).poi(3), item0(it%).data(0).n(0), item0(it%).data(0).n(1), _
                                 item0(it%).data(0).n(2), item0(it%).data(0).n(3), item0(it%).data(0).line_no(0), _
                                  item0(it%).data(0).line_no(1), "1", "1", sqr_string(add_string(general_string(ge%).data(0).value, _
                                      time_string("2", item0(it%).data(0).value, False, False), _
                                        False, False), True, False), temp_record, tn%, 0)
                         If combine_general_string_with_item_value0 > 1 Then
                           Exit Function
                         End If
                         If tv(0) <> "F" And tv(1) <> "F" Then
                           If (m_poi(item0(it%).data(0).poi(0)).data(0).data0.coordinate.X - _
                                m_poi(item0(it%).data(0).poi(1)).data(0).data0.coordinate.X) ^ 2 + _
                              (m_poi(item0(it%).data(0).poi(0)).data(0).data0.coordinate.Y - _
                                m_poi(item0(it%).data(0).poi(1)).data(0).data0.coordinate.Y) ^ 2 > _
                              (m_poi(item0(it%).data(0).poi(2)).data(0).data0.coordinate.X - _
                                m_poi(item0(it%).data(0).poi(3)).data(0).data0.coordinate.X) ^ 2 + _
                              (m_poi(item0(it%).data(0).poi(2)).data(0).data0.coordinate.Y - _
                                m_poi(item0(it%).data(0).poi(3)).data(0).data0.coordinate.Y) ^ 2 Then
                            combine_general_string_with_item_value0 = set_line_value( _
                               item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), tv(0), item0(it%).data(0).n(0), _
                                 item0(it%).data(0).n(1), item0(it%).data(0).line_no(0), temp_record, 0, 0, False)
                            If combine_general_string_with_item_value0 > 1 Then
                              Exit Function
                            End If
                            combine_general_string_with_item_value0 = set_line_value( _
                               item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), tv(1), item0(it%).data(0).n(2), _
                                 item0(it%).data(0).n(3), item0(it%).data(0).line_no(1), temp_record, 0, 0, False)
                            If combine_general_string_with_item_value0 > 1 Then
                              Exit Function
                            End If
                            Else
                            combine_general_string_with_item_value0 = set_line_value( _
                               item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), tv(1), item0(it%).data(0).n(0), _
                                 item0(it%).data(0).n(1), item0(it%).data(0).line_no(0), temp_record, 0, 0, False)
                            If combine_general_string_with_item_value0 > 1 Then
                              Exit Function
                            End If
                            combine_general_string_with_item_value0 = set_line_value( _
                               item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), tv(0), item0(it%).data(0).n(2), _
                                 item0(it%).data(0).n(3), item0(it%).data(0).line_no(1), temp_record, 0, 0, False)
                            If combine_general_string_with_item_value0 > 1 Then
                              Exit Function
                            End If
                           End If
                         End If
                         Else
                         End If

End Function
Public Function combine_eline_with_dpoint_pair(ByVal el%, _
                 ByVal start%, ByVal no_reduce As Byte) As Byte '10.10
Dim i%, k%, l%, no%
Dim n(2) As Integer
Dim m(3) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim last_tn%
Dim dp As point_pair_data0_type
If Deline(el%).record_.no_reduce > 4 Then
 Exit Function
End If
For k% = 0 To 1
n(0) = k%
n(1) = (k% + 1) Mod 2
For l% = 0 To 3
m(0) = l%
m(1) = (l% + 1) Mod 4
m(2) = (l% + 2) Mod 4
m(3) = (l% + 3) Mod 4
dp.poi(2 * m(0)) = Deline(el%).data(0).data0.poi(2 * n(0))
dp.poi(2 * m(0) + 1) = Deline(el%).data(0).data0.poi(2 * n(0) + 1)
dp.poi(2 * m(1)) = -1
Call search_for_point_pair(dp, m(0), n_(0), 1)
dp.poi(2 * m(1)) = 30000
Call search_for_point_pair(dp, m(0), n_(1), 1)   '5.7
If m(0) = 1 Then
m(1) = 0
m(2) = 3
m(3) = 2
ElseIf m(0) = 3 Then
m(1) = 2
m(2) = 1
m(3) = 0
End If
last_tn% = 0
For i% = n_(0) + 1 To n_(1)
no% = Ddpoint_pair(i%).data(0).record.data1.index.i(m(0))
If Ddpoint_pair(no%).record_.no_reduce < 4 And no% > start% Then
'If is_two_record_related(eline_, el%, Deline(el%).data(0).record, _
     dpoint_pair_, no%, Ddpoint_pair(no%).data(0).record) = False Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
'End If
Next i%
For i% = 1 To last_tn%
no% = tn(i%)
combine_eline_with_dpoint_pair = _
 combine_relation_with_dpoint_pair_(eline_, el%, no%, n(0), m(0), no_reduce)
If combine_eline_with_dpoint_pair > 1 Then
 Exit Function
End If
Next i%
Next l%
Next k%
End Function
Public Function combine_eline_with_eline_0(ByVal el%, _
     ByVal no_reduce As Byte) As Byte '10.10
Dim i%, j%, k%, no%, t_n%
Dim n(1) As Integer
Dim m(1) As Integer
Dim n1(3) As Integer
Dim n2(3) As Integer
Dim l1(1) As Integer
Dim l2(1) As Integer
Dim tp(3) As Integer
Dim tp1(3) As Integer
Dim l(1) As Integer
Dim n_(1) As Integer
Dim tn1() As Integer
Dim tn2() As Integer
Dim ty(1) As Byte
Dim last_tn1 As Integer
Dim last_tn2 As Integer
Dim eli As eline_data0_type
Dim temp_record As total_record_type
Dim re As total_record_type
Dim is_no_initial(1) As Integer
Dim c_data(1) As condition_data_type
If Deline(el%).record_.no_reduce > 4 Then
Exit Function
End If
Call add_conditions_to_record(eline_, el%, 0, 0, re.record_data.data0.condition_data)
 re.record_data.data0.theorem_no = 1
 For i% = 0 To 1
 n(0) = i%
  n(1) = (i% + 1) Mod 2
 For j% = 0 To 1
 m(0) = j%
  m(1) = (j% + 1) Mod 2
 eli.line_no(m(0)) = Deline(el%).data(0).data0.line_no(n(0))
 eli.line_no(m(1)) = -1
 Call search_for_eline(eli, m(0) + 2, n_(0), 1) '5.7 原m(0)+2
 eli.line_no(m(1)) = 30000
 Call search_for_eline(eli, m(0) + 2, n_(1), 1)
 last_tn1 = 0
 last_tn2 = 0
 For k% = n_(0) + 1 To n_(1)
  no% = Deline(k%).data(0).record.data1.index.i(m(0) + 2) '5.7
 If no% > 0 And no% < el% And _
   Deline(no%).record_.no_reduce < 4 Then
    If Deline(no%).data(0).data0.line_no(m(1)) = Deline(el%).data(0).data0.line_no(n(1)) Then
     last_tn1 = last_tn1 + 1
      ReDim Preserve tn1(last_tn1) As Integer
       tn1(last_tn1) = no%
    ElseIf line3_value_conclusion = 1 Then
     last_tn2% = last_tn2% + 1
      ReDim Preserve tn2(last_tn2) As Integer
       tn2(last_tn2) = no%
    End If
 End If
 Next k%
 For k% = 1 To last_tn1
   no% = tn1(k%)
   temp_record = re
    Call add_conditions_to_record(eline_, no%, 0, 0, temp_record.record_data.data0.condition_data)
   Call arrange_four_point(Deline(el%).data(0).data0.poi(2 * n(0)), _
            Deline(el%).data(0).data0.poi(2 * n(0) + 1), _
             Deline(no%).data(0).data0.poi(2 * m(0)), _
              Deline(no%).data(0).data0.poi(2 * m(0) + 1), _
            Deline(el%).data(0).data0.n(2 * n(0)), _
             Deline(el%).data(0).data0.n(2 * n(0) + 1), _
              Deline(no%).data(0).data0.n(2 * m(0)), _
               Deline(no%).data(0).data0.n(2 * m(0) + 1), _
                Deline(el%).data(0).data0.line_no(n(0)), _
                 Deline(no%).data(0).data0.line_no(m(0)), _
            tp(0), tp(1), tp(2), tp(3), 0, 0, _
             n1(0), n1(1), n1(2), n1(3), 0, 0, l1(0), l1(1), 0, ty(0), c_data(0), is_no_initial(0))
   Call arrange_four_point(Deline(el%).data(0).data0.poi(2 * n(1)), _
            Deline(el%).data(0).data0.poi(2 * n(1) + 1), _
             Deline(no%).data(0).data0.poi(2 * m(1)), _
              Deline(no%).data(0).data0.poi(2 * m(1) + 1), _
            Deline(el%).data(0).data0.n(2 * n(1)), _
             Deline(el%).data(0).data0.n(2 * n(1) + 1), _
              Deline(no%).data(0).data0.n(2 * m(1)), _
               Deline(no%).data(0).data0.n(2 * m(1) + 1), _
                Deline(el%).data(0).data0.line_no(n(1)), _
                 Deline(no%).data(0).data0.line_no(m(1)), _
                  tp1(0), tp1(1), tp1(2), tp1(3), 0, 0, _
                   n2(0), n2(1), n2(2), n2(3), 0, 0, l2(0), l2(1), 0, ty(1), c_data(1), is_no_initial(1))
   If (ty(0) = 3 Or ty(0) = 5) And (ty(1) = 3 Or ty(1) = 5) Then
     If is_no_initial(0) = 1 Then
      Call add_record_to_record(c_data(0), temp_record.record_data.data0.condition_data)
     ElseIf is_no_initial(1) = 1 Then
      Call add_record_to_record(c_data(1), temp_record.record_data.data0.condition_data)
     End If
     combine_eline_with_eline_0 = set_equal_dline(tp(0), tp(3), _
      tp1(0), tp1(3), n1(0), n1(3), n2(0), n2(3), l1(0), l2(0), _
       0, temp_record, 0, 0, 0, 0, no_reduce, False)
     If combine_eline_with_eline_0 > 1 Then
      Exit Function
     End If
    ElseIf (ty(0) = 4 And ty(1) = 4) _
           Or (ty(0) = 6 And ty(1) = 6) Then
      combine_eline_with_eline_0 = set_equal_dline(tp(0), tp(1), _
      tp1(0), tp1(1), n1(0), n1(1), n2(0), n2(1), l1(0), l2(0), _
       0, temp_record, 0, 0, 0, 0, no_reduce, False)
     If combine_eline_with_eline_0 > 1 Then
      Exit Function
     End If
    ElseIf (ty(0) = 7 And ty(1) = 7) _
             Or (ty(0) = 8 And ty(1) = 8) Then
     combine_eline_with_eline_0 = set_equal_dline(tp(2), tp(3), _
      tp1(2), tp1(3), n1(2), n1(3), n2(2), n2(3), l1(1), l2(1), _
       0, temp_record, 0, 0, 0, 0, no_reduce, False)
     If combine_eline_with_eline_0 > 1 Then
      Exit Function
     End If
    ElseIf (ty(0) = 4 And ty(1) = 8) Or _
            (ty(0) = 6 And ty(1) = 7) Then
     combine_eline_with_eline_0 = set_equal_dline(tp(0), tp(1), _
      tp1(2), tp1(3), n1(0), n1(1), n2(2), n2(3), l1(0), l2(1), _
       0, temp_record, 0, 0, 0, 0, no_reduce, False)
     If combine_eline_with_eline_0 > 1 Then
      Exit Function
     End If
    ElseIf (ty(0) = 8 And ty(1) = 4) Or _
             (ty(0) = 7 And ty(1) = 6) Then
         combine_eline_with_eline_0 = set_equal_dline(tp(2), tp(3), _
      tp1(0), tp1(1), n1(2), n1(3), n2(0), n2(1), l1(1), l2(0), _
       0, temp_record, 0, 0, 0, 0, no_reduce, False)
     If combine_eline_with_eline_0 > 1 Then
      Exit Function
     End If
    End If
  Next k%
  For k% = 1 To last_tn2
   no% = tn2(k%)
   temp_record = re
    Call add_conditions_to_record(eline_, no%, 0, 0, temp_record.record_data.data0.condition_data)
    Call arrange_four_point(Deline(el%).data(0).data0.poi(2 * n(0)), _
            Deline(el%).data(0).data0.poi(2 * n(0) + 1), Deline(no%).data(0).data0.poi(2 * m(0)), _
             Deline(no%).data(0).data0.poi(2 * m(0) + 1), 0, 0, 0, 0, 0, 0, _
              tp(0), tp(1), tp(2), tp(3), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ty(0), c_data(0), is_no_initial(0))
    If ty(0) = 3 Or ty(0) = 5 Then
    If is_no_initial(0) = 1 Then
          Call add_record_to_record(c_data(0), temp_record.record_data.data0.condition_data)
    End If
    combine_eline_with_eline_0 = set_three_line_value(tp(0), tp(3), _
     Deline(el%).data(0).data0.poi(2 * n(1)), Deline(el%).data(0).data0.poi(2 * n(1) + 1), _
      Deline(no%).data(0).data0.poi(2 * m(1)), Deline(no%).data(0).data0.poi(2 * m(1) + 1), _
     n1(0), n1(3), Deline(el%).data(0).data0.n(2 * n(1)), Deline(el%).data(0).data0.n(2 * n(1) + 1), _
      Deline(no%).data(0).data0.n(2 * m(1)), Deline(no%).data(0).data0.n(2 * m(1) + 1), _
     l1(0), Deline(el%).data(0).data0.line_no(n(1)), Deline(no%).data(0).data0.line_no(m(1)), _
       "1", "-1", "-1", "0", temp_record, 0, no_reduce, 1)
      If combine_eline_with_eline_0 > 1 Then
       Exit Function
      End If
    ElseIf ty(0) = 4 Then
    combine_eline_with_eline_0 = set_three_line_value(tp(0), tp(1), _
     Deline(el%).data(0).data0.poi(2 * n(1)), Deline(el%).data(0).data0.poi(2 * n(1) + 1), _
      Deline(no%).data(0).data0.poi(2 * m(1)), Deline(no%).data(0).data0.poi(2 * m(1) + 1), _
     n1(0), n1(1), Deline(el%).data(0).data0.n(2 * n(1)), Deline(el%).data(0).data0.n(2 * n(1) + 1), _
      Deline(no%).data(0).data0.n(2 * m(1)), Deline(no%).data(0).data0.n(2 * m(1) + 1), _
     l1(0), Deline(el%).data(0).data0.line_no(n(1)), Deline(no%).data(0).data0.line_no(2 * m(1)), _
      "1", "-1", "1", "0", temp_record, 0, no_reduce, 1)
      If combine_eline_with_eline_0 > 1 Then
       Exit Function
      End If
    ElseIf ty(0) = 6 Then
    combine_eline_with_eline_0 = set_three_line_value(tp(0), tp(1), _
     Deline(el%).data(0).data0.poi(2 * n(1)), Deline(el%).data(0).data0.poi(2 * n(1) + 1), _
      Deline(no%).data(0).data0.poi(2 * m(1)), Deline(no%).data(0).data0.poi(2 * m(1) + 1), _
     n1(0), n1(1), Deline(el%).data(0).data0.n(2 * n(1)), Deline(el%).data(0).data0.n(2 * n(1) + 1), _
      Deline(no%).data(0).data0.n(2 * m(1)), Deline(no%).data(0).data0.n(2 * m(1) + 1), _
     l1(0), Deline(el%).data(0).data0.line_no(n(1)), Deline(no%).data(0).data0.line_no(m(1)), _
       "1", "1", "-1", "0", temp_record, 0, no_reduce, 1)
      If combine_eline_with_eline_0 > 1 Then
       Exit Function
      End If
    ElseIf ty(0) = 7 Then
    combine_eline_with_eline_0 = set_three_line_value(tp(2), tp(3), _
     Deline(el%).data(0).data0.poi(2 * n(1)), Deline(el%).data(0).data0.poi(2 * n(1) + 1), _
      Deline(no%).data(0).data0.poi(2 * m(1)), Deline(no%).data(0).data0.poi(2 * m(1) + 1), _
    n1(2), n1(3), Deline(el%).data(0).data0.n(2 * n(1)), Deline(el%).data(0).data0.n(2 * n(1) + 1), _
      Deline(no%).data(0).data0.n(2 * m(1)), Deline(no%).data(0).data0.n(2 * m(1) + 1), _
    l1(1), Deline(el%).data(0).data0.line_no(n(1)), Deline(no%).data(0).data0.n(2 * m(1)), _
       "1", "1", "-1", "0", temp_record, 0, no_reduce, 1)
      If combine_eline_with_eline_0 > 1 Then
       Exit Function
      End If
    ElseIf ty(0) = 8 Then
    combine_eline_with_eline_0 = set_three_line_value(tp(2), tp(3), _
     Deline(el%).data(0).data0.poi(2 * n(1)), Deline(el%).data(0).data0.poi(2 * n(1) + 1), _
      Deline(no%).data(0).data0.poi(2 * m(1)), Deline(no%).data(0).data0.poi(2 * m(1) + 1), _
     n1(2), n2(3), Deline(el%).data(0).data0.n(2 * n(1)), Deline(el%).data(0).data0.n(2 * n(1) + 1), _
      Deline(no%).data(0).data0.n(2 * m(1)), Deline(no%).data(0).data0.n(2 * m(1) + 1), _
     l1(1), Deline(el%).data(0).data0.line_no(n(1)), Deline(no%).data(0).data0.line_no(m(1)), _
       "1", "-1", "1", "0", temp_record, 0, no_reduce, 1)
      If combine_eline_with_eline_0 > 1 Then
       Exit Function
      End If
    End If
  Next k%
  Next j%
  Next i%
End Function
Public Function combine_eline_with_eline(ByVal el%, _
      ByVal no_reduce As Byte) As Byte '10.10
Dim k%, l%
If Deline(el%).record_.no_reduce > 4 Then
 Exit Function
End If
For k% = 0 To 1
 For l% = 0 To 1
 combine_eline_with_eline = _
   combine_eline_with_eline0(el%, k%, l%, no_reduce)
  If combine_eline_with_eline > 1 Then
   Exit Function
  End If
 Next l%
Next k%
 combine_eline_with_eline = combine_eline_with_eline_0(el%, no_reduce)
End Function
Public Function combine_relation_with_relation(ByVal re%, _
    ByVal no_reduce As Byte) As Byte
Dim temp_record As total_record_type
Dim temp_record1 As total_record_type
Dim n(2) As Integer
Dim m(2) As Integer
Dim ty, ty1 As Byte
Dim ty3(1) As Byte
Dim t_y(1) As Byte
Dim t_n() As Integer
Dim n_(1) As Integer
Dim last_tn As Integer
Dim v As String
Dim tn() As Integer
Dim rel As relation_data0_type
Dim ts(2) As String
Dim k%, l%, j%, no%
If Drelation(re%).record_.no_reduce > 4 Then
 Exit Function
End If
For j% = 0 To 2
n(0) = j%
 n(1) = (j% + 1) Mod 3
  n(2) = (j% + 2) Mod 3
If Drelation(re%).data(0).data0.poi(2 * n(0)) > 0 And _
       Drelation(re%).data(0).data0.poi(2 * n(0) + 1) > 0 Then
For k% = 0 To 2
m(0) = k%
 m(1) = (k% + 1) Mod 3
  m(2) = (k% + 2) Mod 3
rel.poi(2 * m(0)) = Drelation(re%).data(0).data0.poi(2 * n(0))
 rel.poi(2 * m(0) + 1) = Drelation(re%).data(0).data0.poi(2 * n(0) + 1)
  rel.poi(2 * m(1)) = -1
Call search_for_relation(rel, m(0), n_(0), 1)
rel.poi(2 * m(1)) = 30000
Call search_for_relation(rel, m(0), n_(1), 1)  '5.7
last_tn = 0
For l% = n_(0) + 1 To n_(1)
 no% = Drelation(l%).data(0).record.data1.index.i(m(0))
If no% > 0 And no% < re% And _
    Drelation(no%).record_.no_reduce < 4 Then
'If is_two_record_related(relation_, no%, Drelation(no%).data(0).record, _
     relation_, re%, Drelation(re%).data(0).record) = False Then
last_tn = last_tn + 1
ReDim Preserve tn(last_tn) As Integer
 tn(last_tn) = no%
End If
'End If
Next l%
For l% = 1 To last_tn
no% = tn(l%)
combine_relation_with_relation = combine_relation_with_relation_( _
   relation_, re%, relation_, no%, n(0), m(0))
If combine_relation_with_relation > 1 Then
 Exit Function
End If
Next l%
 Next k%
 End If
Next j%
ts(0) = "1"
Call read_ratio_from_Drelation(re%, j%, ts(1), ts(2))
Call add_conditions_to_record(relation_, re%, 0, 0, temp_record1.record_data.data0.condition_data)
        temp_record1.record_data.data0.theorem_no = 1
rel.value = Drelation(re%).data(0).data0.value
rel.poi(0) = -1
Call search_for_relation(rel, 3, n_(0), 1) '5.7原6
'rel.value = Drelation(re%).data(0).data0.value
rel.poi(0) = 30000
Call search_for_relation(rel, 3, n_(1), 1)
last_tn = 0
For k% = n_(0) + 1 To n_(1)
no% = Drelation(k%).data(0).record.data1.index.i(3)
If no% > 0 And no% < re% And _
     Drelation(no%).record_.no_reduce < 4 Then
last_tn = last_tn + 1
ReDim Preserve t_n(last_tn) As Integer
t_n(last_tn) = no%
End If
Next k%
For k% = 1 To last_tn
no% = t_n(k%)
 temp_record = temp_record1
  Call add_conditions_to_record(relation_, no%, 0, 0, temp_record.record_data.data0.condition_data)
'If Drelation(i%).data(0).data0.value = Drelation(j%).data(0).data0.value Then
  'combine_relation_with_relation = _
   set_dpoint_pair0(Drelation(re%).data(0).data0.poi(0), Drelation(re%).data(0).data0.poi(1), _
   Drelation(re%).data(0).data0.poi(2), Drelation(i%).data(0).data0.poi(3), Drelation(no%).data(0).data0.poi(0), _
    Drelation(no%).data(0).data0.poi(1), Drelation(no%).data(0).data0.poi(2), Drelation(no%).data(0).data0.poitn(3), _
     temp_record, True, 0, 0, no_reduce)
   'If combine_relation_with_relation = 0 Then
    ' Drelation(no%).record_.no_reduce = 2
    '  If combine_relation_with_relation > 1 Then
     '  Exit Function
     ' End If
   'End If
 combine_relation_with_relation = combine_relation_with_same_ratio( _
     Drelation(re%).data(0).data0.poi(0), Drelation(re%).data(0).data0.poi(1), _
       Drelation(re%).data(0).data0.poi(2), Drelation(re%).data(0).data0.poi(3), _
         Drelation(no%).data(0).data0.poi(0), Drelation(no%).data(0).data0.poi(1), _
           Drelation(no%).data(0).data0.poi(2), Drelation(no%).data(0).data0.poi(3), _
     Drelation(re%).data(0).data0.n(0), Drelation(re%).data(0).data0.n(1), _
       Drelation(re%).data(0).data0.n(2), Drelation(re%).data(0).data0.n(3), _
         Drelation(no%).data(0).data0.n(0), Drelation(no%).data(0).data0.n(1), _
           Drelation(no%).data(0).data0.n(2), Drelation(no%).data(0).data0.n(3), _
     Drelation(re%).data(0).data0.line_no(0), Drelation(re%).data(0).data0.line_no(1), _
      Drelation(no%).data(0).data0.line_no(0), Drelation(no%).data(0).data0.line_no(1), _
             Drelation(re%).data(0).data0.value, temp_record.record_data, no_reduce)
 If combine_relation_with_relation > 1 Then
  Exit Function
 End If
 If th_chose(20).chose = 1 Then
 temp_record.record_data.data0.theorem_no = 20
 If Drelation(re%).data(0).data0.poi(4) > 0 And Drelation(re%).data(0).data0.poi(5) > 0 And _
      Drelation(no%).data(0).data0.poi(4) > 0 And Drelation(no%).data(0).data0.poi(5) > 0 Then
  If Drelation(re%).data(0).data0.poi(0) = Drelation(no%).data(0).data0.poi(0) Then
 combine_relation_with_relation = _
   set_similar_triangle(Drelation(no%).data(0).data0.poi(0), _
          Drelation(no%).data(0).data0.poi(1), Drelation(re%).data(0).data0.poi(1), _
           Drelation(no%).data(0).data0.poi(0), Drelation(no%).data(0).data0.poi(3), _
            Drelation(re%).data(0).data0.poi(3), temp_record, 0, no_reduce, 1)
      If combine_relation_with_relation > 1 Then
       Exit Function
      End If
  ElseIf Drelation(re%).data(0).data0.poi(1) = Drelation(no%).data(0).data0.poi(1) Then
 combine_relation_with_relation = _
   set_similar_triangle(Drelation(no%).data(0).data0.poi(1), _
          Drelation(no%).data(0).data0.poi(0), Drelation(re%).data(0).data0.poi(0), _
           Drelation(no%).data(0).data0.poi(1), Drelation(no%).data(0).data0.poi(3), _
            Drelation(re%).data(0).data0.poi(3), temp_record, 0, no_reduce, 1)
      If combine_relation_with_relation > 1 Then
       Exit Function
      End If
  ElseIf Drelation(re%).data(0).data0.poi(3) = Drelation(no%).data(0).data0.poi(3) Then
 combine_relation_with_relation = _
   set_similar_triangle(Drelation(no%).data(0).data0.poi(3), _
          Drelation(no%).data(0).data0.poi(0), Drelation(re%).data(0).data0.poi(0), _
           Drelation(no%).data(0).data0.poi(3), Drelation(no%).data(0).data0.poi(1), _
            Drelation(re%).data(0).data0.poi(1), temp_record, 0, no_reduce, 1)
      If combine_relation_with_relation > 1 Then
       Exit Function
      End If
  End If
 End If
 End If
 Next k%
  v = divide_string("1", Drelation(re%).data(0).data0.value, True, False)
temp_record.record_data.data0.theorem_no = 1
rel.value = v
rel.poi(0) = -1
Call search_for_relation(rel, 3, n_(0), 1) '5.7原6
'rel.value = v
rel.poi(0) = 30000
Call search_for_relation(rel, 3, n_(1), 1)
last_tn = 0
For k% = n_(0) + 1 To n_(1)
no% = Drelation(k%).data(0).record.data1.index.i(3)
If no% > 0 And no% <> re% And _
 Drelation(no%).record_.no_reduce < 4 Then
last_tn = last_tn + 1
ReDim Preserve t_n(last_tn) As Integer
t_n(last_tn) = no%
End If
Next k%
For k% = 1 To last_tn
no% = t_n(k%)
 temp_record = temp_record1
  Call add_conditions_to_record(relation_, no%, 0, 0, temp_record.record_data.data0.condition_data)
 'combine_relation_with_relation = _
  'set_dpoint_pair(Drelation(i%).tn(0), Drelation(i%).tn(1), _
  'Drelation(i%).tn(2), Drelation(i%).tn(3), Drelation(no%).tn(2), _
   'Drelation(no%).tn(3), Drelation(no%).tn(0), Drelation(no%).tn(1), _
    'Drelation(i%).line_no(0), Drelation(i%).line_no(1), _
     'Drelation(no%).line_no(1), Drelation(no%).line_no(0), temp_record, _
      'True, 0, 0, no_reduce)
      'If combine_relation_with_relation > 1 Then
       'Exit Function
      'End If
 combine_relation_with_relation = combine_relation_with_same_ratio( _
     Drelation(re%).data(0).data0.poi(0), Drelation(re%).data(0).data0.poi(1), _
       Drelation(re%).data(0).data0.poi(2), Drelation(re%).data(0).data0.poi(3), _
         Drelation(no%).data(0).data0.poi(2), Drelation(no%).data(0).data0.poi(3), _
           Drelation(no%).data(0).data0.poi(0), Drelation(no%).data(0).data0.poi(1), _
     Drelation(re%).data(0).data0.n(0), Drelation(re%).data(0).data0.n(1), _
       Drelation(re%).data(0).data0.n(2), Drelation(re%).data(0).data0.n(3), _
         Drelation(no%).data(0).data0.n(2), Drelation(no%).data(0).data0.n(3), _
           Drelation(no%).data(0).data0.n(0), Drelation(no%).data(0).data0.n(1), _
     Drelation(re%).data(0).data0.line_no(0), Drelation(re%).data(0).data0.line_no(1), _
      Drelation(no%).data(0).data0.line_no(1), Drelation(no%).data(0).data0.line_no(0), _
             Drelation(re%).data(0).data0.value, temp_record.record_data, no_reduce)
 If combine_relation_with_relation > 1 Then
  Exit Function
 End If
 If th_chose(20).chose = 1 Then
 temp_record.record_data.data0.theorem_no = 52
 If Drelation(re%).data(0).data0.poi(4) > 0 And Drelation(re%).data(0).data0.poi(5) > 0 And _
      Drelation(no%).data(0).data0.poi(4) > 0 And Drelation(no%).data(0).data0.poi(5) > 0 Then
  If Drelation(re%).data(0).data0.poi(3) = Drelation(no%).data(0).data0.poi(0) Then
   combine_relation_with_relation = _
   set_similar_triangle(Drelation(no%).data(0).data0.poi(0), _
          Drelation(no%).data(0).data0.poi(1), Drelation(re%).data(0).data0.poi(1), _
           Drelation(no%).data(0).data0.poi(0), Drelation(no%).data(0).data0.poi(3), _
            Drelation(re%).data(0).data0.poi(0), temp_record, 0, no_reduce, 1)
      If combine_relation_with_relation > 1 Then
       Exit Function
      End If
  ElseIf Drelation(re%).data(0).data0.poi(1) = Drelation(no%).data(0).data0.poi(1) Then
 combine_relation_with_relation = _
   set_similar_triangle(Drelation(no%).data(0).data0.poi(1), _
          Drelation(no%).data(0).data0.poi(0), Drelation(re%).data(0).data0.poi(3), _
           Drelation(no%).data(0).data0.poi(1), Drelation(no%).data(0).data0.poi(3), _
            Drelation(re%).data(0).data0.poi(0), temp_record, 0, no_reduce, 1)
      If combine_relation_with_relation > 1 Then
       Exit Function
      End If
  ElseIf Drelation(re%).data(0).data0.poi(0) = Drelation(no%).data(0).data0.poi(3) Then
 combine_relation_with_relation = _
   set_similar_triangle(Drelation(no%).data(0).data0.poi(3), _
          Drelation(no%).data(0).data0.poi(1), Drelation(re%).data(0).data0.poi(1), _
           Drelation(no%).data(0).data0.poi(3), Drelation(no%).data(0).data0.poi(0), _
            Drelation(re%).data(0).data0.poi(3), temp_record, 0, no_reduce, 1)
      If combine_relation_with_relation > 1 Then
       Exit Function
      End If
  End If
 End If
 End If
Next k%
'**************************************************************************
End Function


Public Function combine_relation_with_line_value(ByVal r%, _
             ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, tn%
Dim num_string As String
'If Drelation(r%).record_.no_reduce = False Then
 For i% = 0 To 2
  If is_line_value(Drelation(r%).data(0).data0.poi(2 * i%), Drelation(r%).data(0).data0.poi(2 * i% + 1), _
              Drelation(r%).data(0).data0.n(2 * i%), Drelation(r%).data(0).data0.n(2 * i% + 1), _
               Drelation(r%).data(0).data0.line_no(i%), "", tn%, -1000, 0, 0, 0, _
                line_value_data0) = 1 Then
  If tn% > start% Then
    If line_value(tn%).record_.no_reduce < 255 Then
    combine_relation_with_line_value = _
      combine_relation_with_relation_(relation_, r%, line_value_, tn%, _
        i%, 0)
  If combine_relation_with_line_value > 1 Then
   Exit Function
  End If
  End If
  End If
  End If
 Next i%
End Function

Public Function combine_relation_with_eline(ByVal r%, _
          ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, no%
Dim n(2) As Integer
Dim m(1) As Integer
Dim tn() As Integer
Dim last_tn%
Dim n_(1) As Integer
Dim el As eline_data0_type
Dim v(1) As String
If Drelation(r%).record_.no_reduce > 4 Then
 Exit Function
End If
For i% = 0 To 2
n(0) = i%
 n(1) = (i% + 1) Mod 3
  n(2) = (i% + 2) Mod 3
For j% = 0 To 1
 m(0) = j%
  m(1) = (j% + 1) Mod 2
el.poi(2 * m(0)) = Drelation(r%).data(0).data0.poi(2 * n(0))
el.poi(2 * m(0) + 1) = Drelation(r%).data(0).data0.poi(2 * n(0) + 1)
el.poi(2 * m(1)) = -1
Call search_for_eline(el, m(0), n_(0), 1)  '5.7
el.poi(2 * m(1)) = 30000
Call search_for_eline(el, m(0), n_(1), 1)
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
 no% = Deline(k%).data(0).record.data1.index.i(m(0))
  If no > start% And Deline(no%).record_.no_reduce < 4 Then
'If is_two_record_related(eline_, no%, Deline(no%).data(0).record, _
      relation_, r%, Drelation(r%).data(0).record) = False Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
'End If
Next k%
For k% = 1 To last_tn%
 no% = tn(k%)
 combine_relation_with_eline = _
   combine_relation_with_relation_(relation_, r%, eline_, no%, n(0), m(0))
 If combine_relation_with_eline > 0 Then
  Exit Function
 End If
Next k%
Next j%
Next i%
End Function

Public Function combine_relation_with_mid_point(ByVal r%, _
          ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, no%
Dim n(2) As Integer
Dim m(5) As Integer
Dim tv(1) As String
Dim tv1(1) As String
Dim md As mid_point_data0_type
For i% = 0 To 2
  n(0) = i%
  n(1) = (i% + 1) Mod 3
  n(2) = (i% + 2) Mod 3
For j% = 0 To 2
 If j% = 0 Then
m(0) = 0
m(1) = 1
m(2) = 1
m(3) = 2
m(4) = 0
m(5) = 2
ElseIf j% = 1 Then
m(0) = 1
m(1) = 2
m(2) = 0
m(3) = 2
m(4) = 0
m(5) = 1
Else
m(0) = 0
m(1) = 2
m(2) = 0
m(3) = 1
m(4) = 1
m(5) = 2
End If
 md.poi(m(0)) = Drelation(r%).data(0).data0.poi(2 * n(0))
 md.poi(m(1)) = Drelation(r%).data(0).data0.poi(2 * n(0) + 1)
If search_for_mid_point(md, j%, no%, 2) Then  '5.7原j%+3
If no% > start% And Dmid_point(no%).record_.no_reduce < 4 Then
'If is_two_record_related(midpoint_, no%, Dmid_point(no%).data(0).record, _
      relation_, r%, Drelation(r%).data(0).record) = False Then
combine_relation_with_mid_point = _
  combine_relation_with_relation_(relation_, r%, midpoint_, no%, _
   n(0), j%)
If combine_relation_with_mid_point > 1 Then
 Exit Function
End If
End If
End If
'End If
Next j%
Next i%
End Function
Public Function combine_mid_point_with_relation(ByVal md%, _
          ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, no%
Dim n(2) As Integer
Dim m(2) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim last_tn%
Dim dr As relation_data0_type
Dim tp(5) As Integer
If Dmid_point(md%).record_.no_reduce > 4 Then
 Exit Function
End If
tp(0) = Dmid_point(md%).data(0).data0.poi(0)
tp(1) = Dmid_point(md%).data(0).data0.poi(1)
tp(2) = Dmid_point(md%).data(0).data0.poi(1)
tp(3) = Dmid_point(md%).data(0).data0.poi(2)
tp(4) = Dmid_point(md%).data(0).data0.poi(0)
tp(5) = Dmid_point(md%).data(0).data0.poi(2)
For i% = 0 To 2
n(0) = i%
n(1) = (i% + 1) Mod 3
n(2) = (i% + 2) Mod 3
For j% = 0 To 2
m(0) = j%
m(1) = (j% + 1) Mod 3
m(2) = (j% + 2) Mod 3
If m(0) < 2 Then
dr.line_no(m(0)) = Dmid_point(md%).data(0).data0.line_no
End If
dr.poi(2 * m(0)) = tp(2 * n(0))
dr.poi(2 * m(0) + 1) = tp(2 * n(0) + 1)
dr.poi(2 * m(1)) = -1
Call search_for_relation(dr, m(0), n_(0), 1)  '5.7
dr.poi(2 * m(1)) = 30000
Call search_for_relation(dr, m(0), n_(1), 1)
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
no% = Drelation(k%).data(0).record.data1.index.i(m(0))
If no% > start% And Drelation(no%).record_.no_reduce < 4 Then
'If is_two_record_related(relation_, no%, Drelation(no%).data(0).record, _
     midpoint_, md%, Dmid_point(md%).data(0).record) = False Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
'End If
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
combine_mid_point_with_relation = _
 combine_relation_with_relation_(midpoint_, md%, relation_, no%, _
  n(0), m(0))
If combine_mid_point_with_relation > 1 Then
 Exit Function
End If
Next k%
Next j%
Next i%
End Function


Public Function combine_mid_point_with_line_value(ByVal m%, _
              ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, tn%
Dim n(5) As Integer
Dim v(1) As String
Dim temp_record As total_record_type
Dim re As total_record_type
If Dmid_point(m%).record_.no_reduce > 4 Then
 Exit Function
End If
Call add_conditions_to_record(midpoint_, m%, 0, 0, re.record_data.data0.condition_data)
   re.record_data.data0.theorem_no = 1
For i% = 0 To 2
If i% = 0 Then
n(0) = 0
n(1) = 1
n(2) = 1
n(3) = 2
n(4) = 0
n(5) = 2
ElseIf i% = 1 Then
n(0) = 1
n(1) = 2
n(2) = 0
n(3) = 2
n(4) = 0
n(5) = 1
Else
n(0) = 0
n(1) = 2
n(2) = 0
n(3) = 1
n(4) = 1
n(5) = 2
End If
If is_line_value(Dmid_point(m%).data(0).data0.poi(n(0)), Dmid_point(m%).data(0).data0.poi(n(1)), _
     Dmid_point(m%).data(0).data0.n(n(0)), Dmid_point(m%).data(0).data0.n(n(1)), _
      Dmid_point(m%).data(0).data0.line_no, "", tn%, -1000, 0, 0, 0, line_value_data0) = 1 Then
If tn% > start% Then
If line_value(tn%).record_.no_reduce < 255 Then
temp_record = re
Call add_conditions_to_record(line_value_, tn%, 0, 0, temp_record.record_data.data0.condition_data)
combine_mid_point_with_line_value = _
  combine_relation_with_relation_(midpoint_, m%, _
    line_value_, tn%, i%, 0)
If combine_mid_point_with_line_value > 1 Then
 Exit Function
End If
End If
End If
End If
Next i%
End Function

Public Function combine_eline_with_line_value(ByVal e%, _
              ByVal start%, ByVal no_reduce As Byte) As Byte '10.10
Dim i%, j%, no%
Dim n(1) As Integer
Dim temp_record As total_record_type
If Deline(e%).record_.no_reduce > 4 Then
 Exit Function
End If
 For i% = 0 To 1
  n(0) = i%
   n(1) = (i% + 1) Mod 2
If is_line_value(Deline(e%).data(0).data0.poi(2 * n(0)), _
     Deline(e%).data(0).data0.poi(2 * n(0) + 1), _
      Deline(e%).data(0).data0.n(2 * n(0)), Deline(e%).data(0).data0.n(2 * n(0) + 1), _
       Deline(e%).data(0).data0.line_no(n(0)), "", no%, -1000, 0, 0, 0, line_value_data0) = 1 Then
If no% > start% Then
 If line_value(no%).record_.no_reduce < 255 Then
 combine_eline_with_line_value = _
   combine_relation_with_relation_(eline_, e%, line_value_, no%, _
    n(0), 0)
 If combine_eline_with_line_value > 1 Then
  Exit Function
 End If
End If
End If
End If
 For j% = 1 To last_conditions.last_cond(1).line_value_no
  If line_value(j%).data(0).data0.line_no = Deline(e%).data(0).data0.line_no(i%) Then
    combine_eline_with_line_value = combine_eline_with_line_value0(e%, j%, i%)
      If combine_eline_with_line_value > 1 Then
       Exit Function
      End If
  End If
 Next j%
Next i%
End Function
Public Function combine_eline_with_midpoint(ByVal e%, _
               ByVal start%, ByVal no_reduce As Byte) As Byte '10.10
Dim tn() As Integer
Dim n(1) As Integer
Dim m(5) As Integer
Dim v(1) As String
Dim md As mid_point_data0_type
Dim temp_record As total_record_type
Dim i%, j%, k%, no%, last_tn%
If Deline(e%).record_.no_reduce > 4 Then
 Exit Function
End If
For i% = 0 To 1
n(0) = i%
 n(1) = (i% + 1) Mod 2
For j% = 0 To 2
If j% = 0 Then
 m(0) = 0
  m(1) = 1
   m(2) = 1
    m(3) = 2
     m(4) = 0
      m(5) = 2
ElseIf j% = 1 Then
 m(0) = 1
  m(1) = 2
   m(2) = 0
    m(3) = 2
     m(4) = 0
      m(5) = 1
Else
 m(0) = 0
  m(1) = 2
   m(2) = 0
    m(3) = 1
     m(4) = 1
      m(5) = 2
End If
md.poi(m(0)) = Deline(e%).data(0).data0.poi(2 * n(0))
md.poi(m(1)) = Deline(e%).data(0).data0.poi(2 * n(0) + 1)
If search_for_mid_point(md, j%, no%, 2) Then  '5.7原j%+3
If no% > start% And Dmid_point(no%).record_.no_reduce < 4 Then
If j% < 2 Then
 temp_record.record_data.data0.condition_data.condition_no = 2
 temp_record.record_data.data0.condition_data.condition(1).ty = eline_
 temp_record.record_data.data0.condition_data.condition(2).ty = midpoint_
 temp_record.record_data.data0.condition_data.condition(1).no = e
 temp_record.record_data.data0.condition_data.condition(2).no = no%
 If n(0) = 0 Then
  If Deline(e%).data(0).data0.poi(2) = Dmid_point(no%).data(0).data0.poi(1) Then
   combine_eline_with_midpoint = set_dverti( _
     line_number0(Dmid_point(no%).data(0).data0.poi(0), Deline(e%).data(0).data0.poi(3), 0, 0), _
      line_number0(Dmid_point(no%).data(0).data0.poi(2), Deline(e%).data(0).data0.poi(3), 0, 0), _
        temp_record, 0, 0, False)
    If combine_eline_with_midpoint > 1 Then
       Exit Function
    End If
  ElseIf Deline(e%).data(0).data0.poi(3) = Dmid_point(no%).data(0).data0.poi(1) Then
   combine_eline_with_midpoint = set_dverti( _
     line_number0(Dmid_point(no%).data(0).data0.poi(0), Deline(e%).data(0).data0.poi(2), 0, 0), _
      line_number0(Dmid_point(no%).data(0).data0.poi(2), Deline(e%).data(0).data0.poi(2), 0, 0), _
        temp_record, 0, 0, False)
    If combine_eline_with_midpoint > 1 Then
       Exit Function
    End If
  End If
 ElseIf n(0) = 1 Then
  If Deline(e%).data(0).data0.poi(0) = Dmid_point(no%).data(0).data0.poi(1) Then
   combine_eline_with_midpoint = set_dverti( _
     line_number0(Dmid_point(no%).data(0).data0.poi(0), Deline(e%).data(0).data0.poi(1), 0, 0), _
      line_number0(Dmid_point(no%).data(0).data0.poi(2), Deline(e%).data(0).data0.poi(1), 0, 0), _
        temp_record, 0, 0, False)
    If combine_eline_with_midpoint > 1 Then
       Exit Function
    End If
  ElseIf Deline(e%).data(0).data0.poi(1) = Dmid_point(no%).data(0).data0.poi(1) Then
   combine_eline_with_midpoint = set_dverti( _
     line_number0(Dmid_point(no%).data(0).data0.poi(0), Deline(e%).data(0).data0.poi(0), 0, 0), _
      line_number0(Dmid_point(no%).data(0).data0.poi(2), Deline(e%).data(0).data0.poi(0), 0, 0), _
        temp_record, 0, 0, False)
    If combine_eline_with_midpoint > 1 Then
       Exit Function
    End If
  End If
 End If
End If
'If is_two_record_related(midpoint_, no%, Dmid_point(no%).data(0).record, _
        eline_, e%, Deline(e%).data(0).record) = False Then
combine_eline_with_midpoint = _
 combine_relation_with_relation_(eline_, e%, midpoint_, no%, _
   n(0), j%)
If combine_eline_with_midpoint > 1 Then
 Exit Function
End If
End If
'End If
End If
Next j%
Next i%
End Function
Public Function combine_eline_with_eline0(ByVal el%, k%, l%, _
       ByVal no_reduce As Byte) As Byte '10.10
Dim i%, no%
Dim n(1) As Integer
Dim m(1) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim e_l  As eline_data0_type
Dim last_tn As Integer
Dim temp_record As total_record_type
Dim re As total_record_type
Call add_conditions_to_record(eline_, el%, 0, 0, re.record_data.data0.condition_data)
re.record_data.data0.theorem_no = 1
n(0) = k%
n(1) = (k% + 1) Mod 2
m(0) = l%
m(1) = (l% + 1) Mod 2
e_l.poi(2 * m(0)) = Deline(el%).data(0).data0.poi(2 * n(0))
e_l.poi(2 * m(0) + 1) = Deline(el%).data(0).data0.poi(2 * n(0) + 1)
e_l.poi(2 * m(1)) = -1
Call search_for_eline(e_l, m(0), n_(0), 1)
e_l.poi(2 * m(1)) = 30000
Call search_for_eline(e_l, m(0), n_(1), 1)  '5.7
last_tn = 0
For i% = n_(0) + 1 To n_(1)
no% = Deline(i%).data(0).record.data1.index.i(m(0))
If no% > 0 And no% < el% And _
    Deline(no%).record_.no_reduce < 4 Then
'If is_two_record_related(eline_, no%, Deline(no%).data(0).record, _
                   eline_, el%, Deline(el%).data(0).record) = False Then
last_tn = last_tn + 1
ReDim Preserve tn(last_tn) As Integer
tn(last_tn) = no%
End If
'End If
Next i%
For i% = 1 To last_tn
no% = tn(i%)
temp_record = re
Call add_conditions_to_record(eline_, no%, 0, 0, temp_record.record_data.data0.condition_data)
combine_eline_with_eline0 = set_equal_dline( _
 Deline(no%).data(0).data0.poi(2 * m(1)), Deline(no%).data(0).data0.poi(2 * m(1) + 1), _
  Deline(el%).data(0).data0.poi(2 * n(1)), Deline(el%).data(0).data0.poi(2 * n(1) + 1), _
   Deline(no%).data(0).data0.n(2 * m(1)), Deline(no%).data(0).data0.n(2 * m(1) + 1), _
    Deline(el%).data(0).data0.n(2 * n(1)), Deline(el%).data(0).data0.n(2 * n(1) + 1), _
     Deline(no%).data(0).data0.line_no(m(1)), Deline(el%).data(0).data0.line_no(n(1)), _
      0, temp_record, 0, 0, 0, 0, 1, False)
If combine_eline_with_eline0 > 1 Then
 Exit Function
End If
If Deline(no%).data(0).data0.eside_tri_point(0) > 0 And _
      Deline(no%).data(0).data0.eside_tri_point(0) = _
        Deline(el%).data(0).data0.eside_tri_point(0) Then
  If Deline(no%).data(0).data0.eside_tri_point(1) = Deline(el%).data(0).data0.eside_tri_point(1) Then
   combine_eline_with_eline0 = set_three_point_on_circle(Deline(no%).data(0).data0.eside_tri_point(1), _
     Deline(no%).data(0).data0.eside_tri_point(2), Deline(el%).data(0).data0.eside_tri_point(2), _
      Deline(no%).data(0).data0.eside_tri_point(0), 0, temp_record)
     If combine_eline_with_eline0 > 1 Then
        Exit Function
     End If
  ElseIf Deline(no%).data(0).data0.eside_tri_point(1) = Deline(el%).data(0).data0.eside_tri_point(2) Then
   combine_eline_with_eline0 = set_three_point_on_circle(Deline(no%).data(0).data0.eside_tri_point(1), _
     Deline(no%).data(0).data0.eside_tri_point(2), Deline(el%).data(0).data0.eside_tri_point(1), _
      Deline(no%).data(0).data0.eside_tri_point(0), 0, temp_record)
     If combine_eline_with_eline0 > 1 Then
        Exit Function
     End If
  ElseIf Deline(no%).data(0).data0.eside_tri_point(2) = Deline(el%).data(0).data0.eside_tri_point(1) Then
   combine_eline_with_eline0 = set_three_point_on_circle(Deline(no%).data(0).data0.eside_tri_point(2), _
     Deline(no%).data(0).data0.eside_tri_point(1), Deline(el%).data(0).data0.eside_tri_point(2), _
      Deline(no%).data(0).data0.eside_tri_point(0), 0, temp_record)
     If combine_eline_with_eline0 > 1 Then
        Exit Function
     End If
  ElseIf Deline(no%).data(0).data0.eside_tri_point(2) = Deline(el%).data(0).data0.eside_tri_point(2) Then
   combine_eline_with_eline0 = set_three_point_on_circle(Deline(no%).data(0).data0.eside_tri_point(2), _
     Deline(no%).data(0).data0.eside_tri_point(1), Deline(el%).data(0).data0.eside_tri_point(1), _
      Deline(no%).data(0).data0.eside_tri_point(0), 0, temp_record)
     If combine_eline_with_eline0 > 1 Then
        Exit Function
     End If
  End If
End If
Next i%
End Function

Public Function combine_relation_with_same_ratio(ByVal p1%, _
   ByVal p2%, ByVal p3%, ByVal p4%, ByVal tp1%, ByVal tp2%, _
    ByVal tp3%, ByVal tp4%, ByVal n1%, ByVal n2%, ByVal n3%, _
      ByVal n4%, ByVal tn1%, ByVal tn2%, ByVal tn3%, ByVal tn4%, _
     ByVal l1%, ByVal l2%, ByVal tl1%, ByVal tl2%, ByVal ratio$, _
      re As record_data_type, ByVal no_reduce As Byte)
Dim tn(2) As Integer
Dim t_y(1) As Byte ' 合比定理
Dim ty_ As Byte
Dim tv$
Dim temp_record As total_record_type
Dim t_p1(1 To 6) As Integer
Dim t_p2(1 To 6) As Integer
Dim t_n1(1 To 6) As Integer
Dim t_n2(1 To 6) As Integer
Dim t_l1(1 To 3) As Integer
Dim t_l2(1 To 3) As Integer
Dim dp As point_pair_data0_type
Dim el As eline_data0_type
Dim is_no_initial(1) As Integer
Dim c_data(1) As condition_data_type
'On Error GoTo combine_relation_with_same_ratio_error
temp_record.record_data = re
temp_record.record_data.data0.theorem_no = 1
Call arrange_four_point(p1%, p2%, p3%, p4%, n1%, n2%, n3%, n4%, _
     l1%, l2%, t_p1(1), t_p1(2), t_p1(3), t_p1(4), t_p1(5), t_p1(6), t_n1(1), t_n1(2), _
      t_n1(3), t_n1(4), t_n1(5), t_n1(6), t_l1(1), t_l1(2), t_l1(3), t_y(0), c_data(0), is_no_initial(0))
Call arrange_four_point(tp1%, tp2%, tp3%, tp4%, tn1%, tn2%, tn3%, tn4%, _
      tl1%, tl2%, t_p2(1), t_p2(2), t_p2(3), t_p2(4), t_p2(5), t_p2(6), t_n2(1), t_n2(2), _
       t_n2(3), t_n2(4), t_n2(5), t_n2(6), t_l2(1), t_l2(2), t_l2(3), t_y(1), c_data(1), is_no_initial(1))
   If (t_y(0) = 3 Or t_y(0) = 5) And (t_y(1) = 3 Or t_y(1) = 5) Then
    If is_equal_dline(t_p1(5), t_p1(6), t_p2(5), t_p2(6), t_n1(5), t_n1(6), t_n2(5), t_n2(6), _
      t_l1(3), t_l2(3), tn(0), -1000, 0, 0, 0, el, tn(1), tn(2), ty_, "", c_data(0)) Then
    Call add_conditions_to_record(ty_, tn(0), tn(1), tn(2), temp_record.record_data.data0.condition_data)
    combine_relation_with_same_ratio = set_equal_dline(p1%, p2%, tp1%, tp2%, _
        n1%, n2%, tn1%, tn2%, l1%, tl1%, 0, temp_record, 0, 0, 0, 0, 0, False)
       If combine_relation_with_same_ratio > 1 Then
          Exit Function
       End If
    combine_relation_with_same_ratio = set_equal_dline(p3%, p4%, tp3%, tp4%, _
        n3%, n4%, tn3%, tn4%, l2%, tl2%, 0, temp_record, 0, 0, 0, 0, 0, False)
       If combine_relation_with_same_ratio > 1 Then
          Exit Function
       End If
    End If
    If t_l1(3) = t_l2(3) Then
    If t_y(0) = 3 And t_y(1) = 5 Then
      If p1% = tp2% Then
         tv$ = divide_string(ratio$, add_string(ratio$, "1", False, False), True, False)
          combine_relation_with_same_ratio = set_Drelation(tp1%, p2%, tp3%, p4%, _
            0, 0, 0, 0, 0, 0, tv$, temp_record, 0, 0, 0, 0, 0, False)
            If combine_relation_with_same_ratio > 1 Then
               Exit Function
            End If
      ElseIf p4% = tp3% Then
         tv$ = divide_string("1", add_string(ratio$, "1", False, False), True, False)
          combine_relation_with_same_ratio = set_Drelation(p3%, tp4%, tp3%, p4%, _
            0, 0, 0, 0, 0, 0, tv$, temp_record, 0, 0, 0, 0, 0, False)
            If combine_relation_with_same_ratio > 1 Then
               Exit Function
            End If
      End If
    ElseIf t_y(0) = 5 And t_y(1) = 3 Then
       If p2% = tp1% Then
         tv$ = divide_string(ratio$, add_string(ratio$, "1", False, False), True, False)
          combine_relation_with_same_ratio = set_Drelation(p1%, tp2%, p3%, tp4%, _
            0, 0, 0, 0, 0, 0, tv$, temp_record, 0, 0, 0, 0, 0, False)
            If combine_relation_with_same_ratio > 1 Then
               Exit Function
            End If
       ElseIf p3% = tp4% Then
         tv$ = divide_string("1", add_string(ratio$, "1", False, False), True, False)
          combine_relation_with_same_ratio = set_Drelation(tp3%, p4%, p3%, tp4%, _
            0, 0, 0, 0, 0, 0, tv$, temp_record, 0, 0, 0, 0, 0, False)
            If combine_relation_with_same_ratio > 1 Then
               Exit Function
            End If
       End If
     End If
    End If
   End If
'************************************************************
Call arrange_four_point(p1%, p2%, tp1%, tp2%, n1%, n2%, tn1%, tn2%, _
     l1%, tl1%, t_p1(1), t_p1(2), t_p2(1), t_p2(2), t_p1(5), t_p1(6), t_n1(1), t_n1(2), _
      t_n2(1), t_n2(2), t_n1(5), t_n1(6), t_l1(1), t_l2(1), t_l1(3), t_y(0), c_data(0), is_no_initial(0))
Call arrange_four_point(p3%, p4%, tp3%, tp4%, n3%, n4%, tn3%, tn4%, _
     l2%, tl2%, t_p1(3), t_p1(4), t_p2(3), t_p2(4), t_p2(5), t_p2(6), t_n1(3), t_n1(4), _
      t_n2(3), t_n2(4), t_n2(5), t_n2(6), t_l1(2), t_l2(2), t_l2(3), t_y(1), c_data(1), is_no_initial(1))
temp_record.record_data = re
temp_record.record_data.data0.theorem_no = 1
   If is_equal_dline(p1%, p2%, tp1%, tp2%, n1%, n2%, n1%, n2%, _
      l1%, tl1%, tn(0), -1000, 0, 0, 0, el, tn(1), tn(2), ty_, "", c_data(0)) Then
    Call add_conditions_to_record(ty_, tn(0), tn(1), tn(2), temp_record.record_data.data0.condition_data)
    combine_relation_with_same_ratio = set_equal_dline(p3%, p4%, tp3%, tp4%, _
        n3%, n4%, tn3%, tn4%, l2%, tl2%, 0, temp_record, 0, 0, 0, 0, 0, False)
       If combine_relation_with_same_ratio > 1 Then
          Exit Function
       End If
   ElseIf is_equal_dline(p3%, p4%, tp3%, tp4%, n3%, n4%, tn3%, tn4%, _
      l2%, tl2%, tn(0), -1000, 0, 0, 0, el, tn(1), tn(2), ty_, "", c_data(0)) Then
    Call add_conditions_to_record(ty_, tn(0), tn(1), tn(2), temp_record.record_data.data0.condition_data)
    combine_relation_with_same_ratio = set_equal_dline(p1%, p2%, tp1%, tp2%, _
        n1%, n2%, tn1%, tn2%, l1%, tl1%, 0, temp_record, 0, 0, 0, 0, 0, False)
       If combine_relation_with_same_ratio > 1 Then
          Exit Function
       End If
   ElseIf (t_y(0) = 3 Or t_y(0) = 5) And (t_y(1) = 3 Or t_y(1) = 5) Then
    If is_equal_dline(t_p1(5), t_p1(6), t_p2(5), t_p2(6), t_n1(5), t_n1(6), t_n2(5), t_n2(6), _
      t_l1(3), t_l2(3), tn(0), -1000, 0, 0, 0, el, tn(1), tn(2), ty_, "", c_data(0)) Then
    Call add_conditions_to_record(ty_, tn(0), tn(1), tn(2), temp_record.record_data.data0.condition_data)
    combine_relation_with_same_ratio = set_equal_dline(p3%, p4%, tp3%, tp4%, _
        n3%, n4%, tn3%, tn4%, l2%, tl2%, 0, temp_record, 0, 0, 0, 0, 0, False)
       If combine_relation_with_same_ratio > 1 Then
          Exit Function
       End If
    combine_relation_with_same_ratio = set_equal_dline(p1%, p2%, tp1%, tp2%, _
        n1%, n2%, tn1%, tn2%, l1%, tl1%, 0, temp_record, 0, 0, 0, 0, 0, False)
       If combine_relation_with_same_ratio > 1 Then
          Exit Function
       End If
    End If
   End If
temp_record.record_data = re
temp_record.record_data.data0.theorem_no = 1
If (t_y(0) = 3 Or t_y(0) = 5) And (t_y(1) = 3 Or t_y(1) = 5) Then
 If is_no_initial(0) = 1 Then
 Call add_record_to_record(c_data(0), temp_record.record_data.data0.condition_data)
 ElseIf is_no_initial(1) = 1 Then
 Call add_record_to_record(c_data(1), temp_record.record_data.data0.condition_data)
 End If
combine_relation_with_same_ratio = _
 set_Drelation(t_p1(5), t_p1(6), t_p2(5), t_p2(6), t_n1(5), t_n1(6), _
     t_n2(5), t_n2(6), t_l1(3), t_l2(3), ratio, temp_record, 0, 0, 0, 0, no_reduce, False)
      If combine_relation_with_same_ratio > 1 Then
       Exit Function
      End If
ElseIf (t_y(0) = 4 And t_y(1) = 4) Or _
         (t_y(0) = 6 And t_y(1) = 6) Then
combine_relation_with_same_ratio = _
 set_Drelation(t_p1(1), t_p1(2), t_p1(3), t_p1(4), t_n1(1), t_n1(2), _
      t_n1(3), t_n1(4), t_l1(1), t_l1(2), ratio, temp_record, 0, 0, 0, 0, no_reduce, False)
      If combine_relation_with_same_ratio > 1 Then
       Exit Function
      End If
ElseIf (t_y(0) = 8 And t_y(1) = 8) Or _
         (t_y(0) = 7 And t_y(1) = 7) Then
combine_relation_with_same_ratio = _
 set_Drelation(t_p2(1), t_p2(2), t_p2(3), t_p2(4), t_n2(1), t_n2(2), _
      t_n2(3), t_n2(4), t_l2(1), t_l2(2), ratio, temp_record, 0, 0, 0, 0, no_reduce, False)
      If combine_relation_with_same_ratio > 1 Then
       Exit Function
      End If
ElseIf (t_y(0) = 4 And t_y(1) = 8) Or _
         (t_y(0) = 6 And t_y(1) = 7) Then
combine_relation_with_same_ratio = _
 set_Drelation(t_p1(1), t_p1(2), t_p2(3), t_p2(4), t_n1(1), t_n1(2), _
      t_n2(3), t_n2(4), t_l1(1), t_l2(2), ratio, temp_record, 0, 0, 0, 0, no_reduce, False)
      If combine_relation_with_same_ratio > 1 Then
       Exit Function
      End If
ElseIf (t_y(0) = 8 And t_y(1) = 4) Or _
         (t_y(0) = 7 And t_y(1) = 8) Then
combine_relation_with_same_ratio = _
 set_Drelation(t_p2(1), t_p2(2), t_p1(3), t_p1(4), t_n2(1), t_n2(2), _
      t_n1(3), t_n1(4), t_l2(1), t_l1(2), ratio, temp_record, 0, 0, 0, 0, no_reduce, False)
      If combine_relation_with_same_ratio > 1 Then
       Exit Function
      End If
End If
If l1% > tl2% Then
Call exchange_two_integer(l1%, tl2%)
Call exchange_two_integer(p1%, tp3%)
Call exchange_two_integer(p2%, tp4%)
Call exchange_two_integer(n1%, tn3%)
Call exchange_two_integer(n2%, tn4%)
End If
If tl1% > l2% Then
Call exchange_two_integer(l2%, tl1%)
Call exchange_two_integer(p3%, tp1%)
Call exchange_two_integer(p4%, tp2%)
Call exchange_two_integer(n3%, tn1%)
Call exchange_two_integer(n4%, tn2%)
End If
If l1% > tl1% Then
Call exchange_two_integer(l1%, tl1%)
Call exchange_two_integer(p1%, tp1%)
Call exchange_two_integer(p2%, tp2%)
Call exchange_two_integer(n1%, tn1%)
Call exchange_two_integer(n2%, tn2%)
Call exchange_two_integer(l2%, tl2%)
Call exchange_two_integer(p3%, tp3%)
Call exchange_two_integer(p4%, tp4%)
Call exchange_two_integer(n3%, tn3%)
Call exchange_two_integer(n4%, tn4%)
End If
'*******
   dp.line_no(0) = l1%
   dp.line_no(1) = tl1%
   dp.line_no(2) = l2%
   dp.line_no(3) = tl2%
   dp.poi(0) = p1%
   dp.poi(1) = p2%
   dp.poi(2) = tp1%
   dp.poi(3) = tp2%
   dp.poi(4) = p3%
   dp.poi(5) = p4%
   dp.poi(6) = tp3%
   dp.poi(7) = tp4%
   dp.n(0) = n1%
   dp.n(1) = n2%
   dp.n(2) = tn1%
   dp.n(3) = tn2%
   dp.n(4) = n3%
   dp.n(5) = n4%
   dp.n(6) = tn3%
   dp.n(7) = tn4%
 temp_record.record_data.data0.theorem_no = 0
 combine_relation_with_same_ratio = set_property_of_dpoint_pair( _
                 dp, 0, 0, temp_record, no_reduce)
 If combine_relation_with_same_ratio > 1 Then
  Exit Function
 End If
combine_relation_with_same_ratio_error:
End Function
Public Function combine_relation_with_dpoint_pair(ByVal re%, _
         ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, k%, l%, no%
Dim n(2) As Integer
Dim m(3) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim v(1) As String
Dim last_tn%
Dim dp As point_pair_data0_type
If Drelation(re%).record_.no_reduce > 4 Then
 Exit Function
End If
For k% = 0 To 2
n(0) = k%
n(1) = (k% + 1) Mod 3
n(2) = (k% + 2) Mod 3
For l% = 0 To 3
m(0) = l%
m(1) = (l% + 1) Mod 4
m(2) = (l% + 2) Mod 4
m(3) = (l% + 3) Mod 4
dp.poi(2 * m(0)) = Drelation(re%).data(0).data0.poi(2 * n(0))
dp.poi(2 * m(0) + 1) = Drelation(re%).data(0).data0.poi(2 * n(0) + 1)
dp.poi(2 * m(1)) = -1
Call search_for_point_pair(dp, m(0), n_(0), 1)
dp.poi(2 * m(1)) = 30000
Call search_for_point_pair(dp, m(0), n_(1), 1)   '5.7
last_tn% = 0
For i% = n_(0) + 1 To n_(1)
no% = Ddpoint_pair(i%).data(0).record.data1.index.i(m(0))
If Ddpoint_pair(no%).record_.no_reduce < 4 And _
  no% > start% Then
'If is_two_record_related(dpoint_pair_, no%, Ddpoint_pair(no%).data(0).record, _
      relation_, re%, Drelation(re%).data(0).record) = False Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
'End If
Next i%
For i% = 1 To last_tn%
no% = tn(i%)
combine_relation_with_dpoint_pair = _
  combine_relation_with_dpoint_pair_(relation_, re%, no%, n(0), m(0), no_reduce)
If combine_relation_with_dpoint_pair > 1 Then
Exit Function
End If
Next i%
Next l%
Next k%
End Function
Public Function combine_mid_point_with_dpoint_pair(ByVal md%, _
          ByVal start%, no_reduce As Byte) As Byte
Dim i%, k%, l%, no%
Dim tp(5) As Integer
Dim n(2) As Integer
Dim m(3) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim v(1) As String
'Dim v1(1) As String
'Dim v2(1) As String
Dim last_tn%
Dim dp As point_pair_data0_type
If Dmid_point(md%).record_.no_reduce > 4 Then
 Exit Function
End If
tp(0) = Dmid_point(md%).data(0).data0.poi(0)
tp(1) = Dmid_point(md%).data(0).data0.poi(1)
tp(2) = Dmid_point(md%).data(0).data0.poi(1)
tp(3) = Dmid_point(md%).data(0).data0.poi(2)
tp(4) = Dmid_point(md%).data(0).data0.poi(0)
tp(5) = Dmid_point(md%).data(0).data0.poi(2)
For k% = 0 To 2
n(0) = k%
 n(1) = (k% + 1) Mod 3
  n(2) = (k% + 2) Mod 3
Call read_ratio_from_relation("1", k%, v(0), v(1), True, 3)
For l% = 0 To 3
m(0) = l%
m(1) = (l% + 1) Mod 4
m(2) = (l% + 2) Mod 4
m(3) = (l% + 3) Mod 4
dp.poi(2 * m(0)) = tp(2 * n(0))
dp.poi(2 * m(0) + 1) = tp(2 * n(0) + 1)
dp.poi(2 * m(1)) = -1
Call search_for_point_pair(dp, m(0), n_(0), 1)
dp.poi(2 * m(1)) = 30000
Call search_for_point_pair(dp, m(0), n_(1), 1)   '5.7
If m(0) = 1 Then
m(1) = 0
m(2) = 3
m(3) = 2
ElseIf m(0) = 3 Then
m(1) = 2
m(2) = 1
m(3) = 0
End If
last_tn% = 0
For i% = n_(0) + 1 To n_(1)
no% = Ddpoint_pair(i%).data(0).record.data1.index.i(m(0))
If Ddpoint_pair(no%).record_.no_reduce < 4 And _
  no% > start% Then
'If is_two_record_related(midpoint_, md%, Dmid_point(md%).data(0).record, _
     dpoint_pair_, no%, Ddpoint_pair(no%).data(0).record) = False Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
'End If
Next i%
For i% = 1 To last_tn%
no% = tn(i%)
combine_mid_point_with_dpoint_pair = _
   combine_relation_with_dpoint_pair_(midpoint_, md%, _
     no%, k%, l%, no_reduce)
 If combine_mid_point_with_dpoint_pair > 1 Then
  Exit Function
 End If
Next i%
Next l%
Next k%
End Function

Public Function combine_eline_with_item(el%, no_reduce As Byte) As Byte '10.10
Dim i%, j%, k%, no%, tn1%
Dim n_(1) As Integer
Dim n(1) As Integer
Dim m(2) As Integer
Dim it As item0_data_type
Dim tn() As Integer
Dim last_tn%
For i% = 0 To 1
 n(0) = i%
  n(1) = (i% + 1) Mod 2
For j% = 0 To 2
 m(0) = j%
  m(1) = (j% + 1) Mod 3
   m(2) = (j% + 2) Mod 3
it.poi(2 * m(0)) = Deline(el%).data(0).data0.poi(2 * n(0))
it.poi(2 * m(0) + 1) = Deline(el%).data(0).data0.poi(2 * n(0) + 1)
it.poi(2 * m(1)) = -1
Call search_for_item0(it, m(0), n_(0), 1)  '5.7
it.poi(2 * m(1)) = 30000
Call search_for_item0(it, m(0), n_(1), 1)
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
 no% = item0(k%).data(0).index(m(0))
 If no% > 0 Then
     last_tn% = last_tn% + 1
     ReDim Preserve tn(last_tn%) As Integer
      tn(last_tn%) = no%
 End If
Next k%
For k% = 1 To last_tn%
 no% = tn(k%)
     combine_eline_with_item = _
      combine_relation_with_item_(eline_, el%, no%, i%, j%, no_reduce)
     If combine_eline_with_item > 1 Then
      Exit Function
     End If
Next k%
Next j%
Next i%

End Function

Public Function combine_item_with_eline(I0 As Integer, no_reduce As Byte) As Byte
Dim i%, j%, no%, k%, tn1%
Dim n(2) As Integer
Dim m(1) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim last_tn%
Dim el As eline_data0_type
For i% = 0 To 2
 n(0) = i%
  n(1) = (i% + 1) Mod 3
   n(2) = (i% + 2) Mod 3
 'If item0(I0).data(0).poi(2 * n(2)) > 0 And item0(I0).data(0).poi(2 * n(2) + 1) > 0 Then
 For j% = 0 To 1
   m(0) = j%
    m(1) = (j% + 1) Mod 2
 el.poi(2 * m(0)) = item0(I0).data(0).poi(2 * n(0))
 el.poi(2 * m(0) + 1) = item0(I0).data(0).poi(2 * n(0) + 1)
 el.poi(2 * m(1)) = -1
 Call search_for_eline(el, m(0), n_(0), 1)
 el.poi(2 * m(1)) = 30000
 Call search_for_eline(el, m(0), n_(1), 1)  '5.7
 last_tn% = 0
 For k% = n_(0) + 1 To n_(1)
 no% = Deline(k%).data(0).record.data1.index.i(m(0))
 If no% > 0 Then
  If Deline(no%).record_.no_reduce < 255 Then
 last_tn% = last_tn% + 1
 ReDim Preserve tn(last_tn%) As Integer
 tn(last_tn%) = no%
 End If
 End If
 Next k%
 For k% = 1 To last_tn%
 no% = tn(k%)
     combine_item_with_eline = _
      combine_relation_with_item_(eline_, no%, I0, j%, i%, no_reduce)
     If combine_item_with_eline > 1 Then
      Exit Function
     End If
Next k%
Next j%
'End If
Next i%
End Function

Public Function combine_item_value_with_general_string(ByVal it%) As Byte
Dim i%, j%, k%, no%
Dim n_(1) As Integer
Dim m(3) As Integer
Dim tn() As Integer
Dim last_tn%
Dim is_zero As Byte
Dim ite(3) As Integer
Dim tp(3) As String
Dim ge As general_string_data_type
Dim temp_record As total_record_type
For i% = 0 To 3
m(0) = i%
m(1) = (i% + 1) Mod 4
m(2) = (i% + 2) Mod 4
m(3) = (i% + 3) Mod 4
ge.item(m(0)) = it%
ge.item(m(1)) = -1
Call search_for_general_string(ge, m(0), n_(0), 1)
ge.item(m(1)) = 30000
Call search_for_general_string(ge, m(0), n_(1), 1)
For j% = n_(0) + 1 To n_(1)
no% = general_string(j%).data(0).record.data1.index.i(m(0))
If no% > 0 And general_string(no%).record_.no_reduce < 4 Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
 tn(last_tn%) = no%
End If
Next j%
last_tn% = 0
For j% = 1 To last_tn%
 no% = tn(j%)
 temp_record.record_data.data0.condition_data.condition_no = 1
  temp_record.record_data.data0.condition_data.condition(1).ty = general_string_
    temp_record.record_data.data0.condition_data.condition(1).no = no%
     temp_record.record_.conclusion_no = general_string(no%).record_.conclusion_no
     temp_record.record_.conclusion_ty = general_string(no%).record_.conclusion_ty
 Call add_record_to_record( _
     item0(it%).data(0).record_for_value.data0.condition_data, _
                            temp_record.record_data.data0.condition_data)
 temp_record.record_data.data0.theorem_no = 1
  For k% = 0 To 3
   If general_string(no%).data(0).item(k%) = it% Then
    ite(k%) = 0
     tp(k%) = time_string(general_string(no%).data(0).para(k%), _
       item0(it%).data(0).value, True, False)
   Else
    ite(k%) = general_string(no%).data(0).item(k%)
     tp(k%) = general_string(no%).data(0).para(k%)
   End If
  Next k%
combine_item_value_with_general_string = _
 set_general_string(ite(0), ite(1), ite(2), ite(3), _
  tp(0), tp(1), tp(2), tp(3), general_string(no%).data(0).value, _
   general_string(no%).record_.conclusion_no, 0, _
    is_zero, temp_record, 0, 0)
 If combine_item_value_with_general_string > 0 Then
  If general_string(no%).data(0).value <> "" Then
     Call set_level_(general_string(no%).record_.no_reduce, 4)
  End If
  If combine_item_value_with_general_string > 1 Then
   Exit Function
  End If
 End If
Next j%
Next i%
End Function
Public Function combine_two_two_line_(t%, _
    ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, no%
Dim n(1) As Integer
Dim m(1) As Integer
Dim tn() As Integer
Dim last_tn%
Dim n_(1) As Integer
Dim t_line As two_line_value_data0_type
Dim temp_record As total_record_type
If two_line_value(t%).record_.no_reduce > 4 Then
 Exit Function
End If
temp_record.record_data.data0.condition_data.condition_no = 2
 temp_record.record_data.data0.condition_data.condition(1).ty = two_line_value_
  temp_record.record_data.data0.condition_data.condition(2).ty = two_line_value_
   temp_record.record_data.data0.condition_data.condition(1).no = t%
    temp_record.record_data.data0.theorem_no = 1
For i% = 0 To 1
 n(0) = i%
  n(1) = (i% + 1) Mod 2
For j% = 0 To 1
  m(0) = j%
   m(1) = (j% + 1) Mod 2
t_line.poi(2 * m(0)) = two_line_value(t%).data(0).data0.poi(2 * n(0))
t_line.poi(2 * m(0) + 1) = two_line_value(t%).data(0).data0.poi(2 * n(0) + 1)
t_line.poi(2 * m(1)) = -1
Call search_for_two_line_value(t_line, m(0), n_(0), 1)
t_line.poi(2 * m(1)) = 30000
Call search_for_two_line_value(t_line, m(0), n_(1), 1)
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
no% = two_line_value(k%).data(0).record.data1.index.i(m(0))
If no% < t% And two_line_value(no%).record_.no_reduce < 4 Then
'If is_two_record_related(two_line_value_, no%, two_line_value(no%).data(0).record, _
    two_line_value_, t%, two_line_value(t%).data(0).record) = False Then
 last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
'End If
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
temp_record.record_data.data0.condition_data.condition(2).no = no%
combine_two_two_line_ = set_two_line_value( _
  two_line_value(t%).data(0).data0.poi(2 * n(1)), two_line_value(t%).data(0).data0.poi(2 * n(1) + 1), _
   two_line_value(no%).data(0).data0.poi(2 * m(1)), two_line_value(no%).data(0).data0.poi(2 * m(1) + 1), _
  two_line_value(t%).data(0).data0.n(2 * n(1)), two_line_value(t%).data(0).data0.n(2 * n(1) + 1), _
   two_line_value(no%).data(0).data0.n(2 * m(1)), two_line_value(no%).data(0).data0.n(2 * m(1) + 1), _
  two_line_value(t%).data(0).data0.line_no(n(1)), two_line_value(no%).data(0).data0.line_no(m(1)), _
   time_string(two_line_value(t%).data(0).data0.para(n(1)), two_line_value(no%).data(0).data0.para(m(0)), True, False), _
     time_string("-1", time_string(two_line_value(no%).data(0).data0.para(m(1)), _
      two_line_value(t%).data(0).data0.para(n(0)), False, False), True, False), minus_string( _
       time_string(two_line_value(t%).data(0).data0.value, two_line_value(no%).data(0).data0.para(m(0)), False, False), _
        time_string(two_line_value(no%).data(0).data0.value, two_line_value(t%).data(0).data0.para(n(0)), False, False), True, False), _
         temp_record, 0, no_reduce)
 If combine_two_two_line_ > 0 Then
  Call set_level_(two_line_value(no%).record_.no_reduce, 4)
 If combine_two_two_line_ > 1 Then
  Exit Function
 End If
 End If
Next k%
Next j%
Next i%
If minus_string(two_line_value(t%).data(0).data0.para(0), "1", True, False) = "0" And _
     minus_string(two_line_value(t%).data(0).data0.para(1), "1", True, False) = "0" Then
 For i% = 1 To t% - 1
  If minus_string(two_line_value(i%).data(0).data0.para(0), "1", True, False) = "0" And _
      minus_string(two_line_value(i%).data(0).data0.para(1), "1", True, False) = "0" Then
   If two_line_value(t%).data(0).data0.line_no(0) = two_line_value(i%).data(0).data0.line_no(0) Then
    combine_two_two_line_ = combine_two_two_line0(t%, 0, i%, 0)
    If combine_two_two_line_ > 1 Then
       Exit Function
    End If
   ElseIf two_line_value(t%).data(0).data0.line_no(0) = two_line_value(i%).data(0).data0.line_no(1) Then
    combine_two_two_line_ = combine_two_two_line0(t%, 0, i%, 1)
    If combine_two_two_line_ > 1 Then
       Exit Function
    End If
   ElseIf two_line_value(t%).data(0).data0.line_no(1) = two_line_value(i%).data(0).data0.line_no(0) Then
    combine_two_two_line_ = combine_two_two_line0(t%, 1, i%, 0)
    If combine_two_two_line_ > 1 Then
       Exit Function
    End If
   ElseIf two_line_value(t%).data(0).data0.line_no(1) = two_line_value(i%).data(0).data0.line_no(1) Then
    combine_two_two_line_ = combine_two_two_line0(t%, 1, i%, 1)
    If combine_two_two_line_ > 1 Then
       Exit Function
    End If
   End If
  End If
 Next i%
End If
End Function
Public Function combine_two_two_line0(ByVal t1%, k%, ByVal t2%, l%) As Byte
Dim temp_record As total_record_type
Dim tp(1) As Integer
Dim tn(1) As Integer
Dim lin As Integer
If two_line_value(t1%).data(0).data0.poi(2 * k% + 1) = two_line_value(t2%).data(0).data0.poi(2 * l%) Then
 tp(0) = two_line_value(t1%).data(0).data0.poi(2 * k%)
 tp(1) = two_line_value(t2%).data(0).data0.poi(2 * l% + 1)
 tn(0) = two_line_value(t1%).data(0).data0.n(2 * k%)
 tn(1) = two_line_value(t2%).data(0).data0.n(2 * l% + 1)
 lin = two_line_value(t1%).data(0).data0.line_no(k%)
ElseIf two_line_value(t1%).data(0).data0.poi(2 * k%) = two_line_value(t2%).data(0).data0.poi(2 * l% + 1) Then
 tp(1) = two_line_value(t1%).data(0).data0.poi(2 * k% + 1)
 tp(0) = two_line_value(t2%).data(0).data0.poi(2 * l%)
 tn(1) = two_line_value(t1%).data(0).data0.n(2 * k% + 1)
 tn(0) = two_line_value(t2%).data(0).data0.n(2 * l%)
 lin = two_line_value(t1%).data(0).data0.line_no(k%)
Else
 Exit Function
End If
Call add_conditions_to_record(two_line_value_, t1%, t2%, 0, temp_record.record_data.data0.condition_data)
 k% = (k% + 1) Mod 2
  l% = (l% + 1) Mod 2
combine_two_two_line0 = set_three_line_value(tp(0), tp(1), two_line_value(t1%).data(0).data0.poi(2 * k%), _
    two_line_value(t1%).data(0).data0.poi(2 * k% + 1), two_line_value(t2%).data(0).data0.poi(2 * l%), _
      two_line_value(t2%).data(0).data0.poi(2 * l% + 1), tn(0), tn(1), _
        two_line_value(t1%).data(0).data0.n(2 * k%), two_line_value(t1%).data(0).data0.n(2 * k% + 1), _
         two_line_value(t2%).data(0).data0.n(2 * l%), two_line_value(t2%).data(0).data0.n(2 * l% + 1), _
          lin, two_line_value(t1%).data(0).data0.line_no(k%), two_line_value(t2%).data(0).data0.line_no(l%), _
            "1", "1", "1", add_string(two_line_value(t1%).data(0).data0.value, two_line_value(t2%).data(0).data0.value, _
               True, False), temp_record, 0, 0, 0)
End Function
Public Function combine_mid_point_with_two_line(md%, _
           ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, no%
Dim n(5) As Integer
Dim m(1) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim t_l As two_line_value_data0_type
Dim last_tn As Integer
If Dmid_point(md%).record_.no_reduce > 4 Then
 Exit Function
End If
For i% = 0 To 2
If i% = 0 Then
n(0) = 0
n(1) = 1
n(2) = 1
n(3) = 2
n(4) = 0
n(5) = 2
ElseIf i% = 1 Then
n(0) = 1
n(1) = 2
n(2) = 0
n(3) = 2
n(4) = 0
n(5) = 1
Else
n(0) = 0
n(1) = 2
n(2) = 0
n(3) = 1
n(4) = 1
n(5) = 2
End If
For j% = 0 To 1
m(0) = j%
m(1) = (j% + 1) Mod 2
t_l.poi(2 * m(0)) = Dmid_point(md%).data(0).data0.poi(n(0))
t_l.poi(2 * m(0) + 1) = Dmid_point(md%).data(0).data0.poi(n(1))
t_l.poi(2 * m(1)) = -1
Call search_for_two_line_value(t_l, j%, n_(0), 1)
t_l.poi(2 * m(1)) = 30000
Call search_for_two_line_value(t_l, j%, n_(1), 1)  '5.7
last_tn = 0
For k% = n_(0) + 1 To n_(1)
no% = two_line_value(k%).data(0).record.data1.index.i(j%)
If two_line_value(no%).record_.no_reduce < 4 And _
    no% > start% Then
    'If is_two_record_related(two_line_value_, no%, two_line_value(no%).data(0).record, _
       midpoint_, md%, Dmid_point(md%).data(0).record) = False Then
last_tn = last_tn + 1
ReDim Preserve tn(last_tn) As Integer
tn(last_tn) = no%
End If
'End If
Next k%
For k% = 1 To last_tn
no% = tn(k%)
combine_mid_point_with_two_line = _
  combine_relation_with_three_line_(midpoint_, md%, _
    two_line_value_, no%, i%, j%)
 If combine_mid_point_with_two_line > 1 Then
  Exit Function
 End If
 Next k%
Next j%
Next i%
End Function

Public Function combine_line_value_with_mid_point(ByVal lv%, _
          ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, tn%, no%
Dim m(5) As Integer
Dim md As mid_point_data0_type
Dim v(1) As String
Dim temp_record As total_record_type
temp_record.record_data.data0.condition_data.condition_no = 2
temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
temp_record.record_data.data0.condition_data.condition(2).ty = midpoint_
temp_record.record_data.data0.condition_data.condition(1).no = lv%
temp_record.record_data.data0.theorem_no = 1
For i% = 0 To 2
Call read_ratio_from_relation("1", i%, v(0), v(1), True, 3)
If i% = 0 Then
m(0) = 0
m(1) = 1
m(2) = 1
m(3) = 2
m(4) = 0
m(5) = 2
ElseIf i% = 1 Then
m(0) = 1
m(1) = 2
m(2) = 0
m(3) = 2
m(4) = 0
m(5) = 1
Else
m(0) = 0
m(1) = 2
m(2) = 0
m(3) = 1
m(4) = 1
m(5) = 2
End If
md.poi(m(0)) = line_value(lv%).data(0).data0.poi(0)
md.poi(m(1)) = line_value(lv%).data(0).data0.poi(1)
If search_for_mid_point(md, i%, no%, 2) Then  '5.7原i%+3
If no% > start% And Dmid_point(no%).record_.no_reduce < 4 Then
 temp_record.record_data.data0.condition_data.condition(2).no = no%
combine_line_value_with_mid_point = set_line_value( _
 Dmid_point(no%).data(0).data0.poi(m(2)), Dmid_point(no%).data(0).data0.poi(m(3)), _
  divide_string(line_value(lv%).data(0).data0.value, v(0), True, False), _
    Dmid_point(no%).data(0).data0.n(m(2)), Dmid_point(no%).data(0).data0.n(m(3)), _
     Dmid_point(no%).data(0).data0.line_no, temp_record, 0, no_reduce, False)
  Call set_level_(Dmid_point(no%).record_.no_reduce, 4)
 If combine_line_value_with_mid_point > 1 Then
  Exit Function
 End If
combine_line_value_with_mid_point = set_line_value( _
 Dmid_point(no%).data(0).data0.poi(m(4)), Dmid_point(no%).data(0).data0.poi(m(5)), _
  divide_string(line_value(lv%).data(0).data0.value, v(1), True, False), _
    Dmid_point(no%).data(0).data0.n(m(4)), Dmid_point(no%).data(0).data0.n(m(5)), _
     Dmid_point(no%).data(0).data0.line_no, temp_record, 0, no_reduce, False)
  Call set_level_(Dmid_point(no%).record_.no_reduce, 4)
 If combine_line_value_with_mid_point > 1 Then
  Exit Function
 End If
End If
End If
Next i%

End Function

Public Function combine_line_value_with_relation(ByVal lv%, _
         ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, k%, no%
Dim n_(1) As Integer
Dim m(2) As Integer
Dim tn() As Integer
Dim last_tn%
Dim re As relation_data0_type
Dim v(1) As String
Dim temp_record As total_record_type
temp_record.record_data.data0.condition_data.condition_no = 2
temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
temp_record.record_data.data0.condition_data.condition(2).ty = relation_
temp_record.record_data.data0.condition_data.condition(1).no = lv%
temp_record.record_data.data0.theorem_no = 1
For i% = 0 To 2
m(0) = i%
m(1) = (i% + 1) Mod 3
m(2) = (i% + 2) Mod 3
re.poi(2 * m(0)) = line_value(lv%).data(0).data0.poi(0)
re.poi(2 * m(0) + 1) = line_value(lv%).data(0).data0.poi(1)
re.poi(2 * m(1)) = -1
Call search_for_relation(re, m(0), n_(0), 1)
re.poi(2 * m(1)) = 30000
Call search_for_relation(re, m(0), n_(1), 1)  '5.7
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
no% = Drelation(k%).data(0).record.data1.index.i(m(0))
If no% > start% And Drelation(no%).record_.no_reduce < 4 Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
combine_line_value_with_relation = _
 combine_relation_with_relation_(relation_, no%, line_value_, _
  lv%, i%, 0)
If combine_line_value_with_relation > 1 Then
  Exit Function
End If
Next k%
Next i%
End Function

Public Function combine_three_line_with_two_line(ByVal t%, _
            ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, no%, last_tn%
Dim n(2) As Integer
Dim m(1) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim t_l As two_line_value_data0_type
For i% = 0 To 2
n(0) = i%
n(1) = (i% + 1) Mod 3
n(2) = (i% + 2) Mod 3
For j% = 0 To 1
m(0) = j%
m(1) = (j% + 1) Mod 2
t_l.poi(2 * m(0)) = line3_value(t%).data(0).data0.poi(2 * n(0))
t_l.poi(2 * m(0) + 1) = line3_value(t%).data(0).data0.poi(2 * n(0) + 1)
t_l.poi(2 * m(1)) = -1
Call search_for_two_line_value(t_l, m(0), n_(0), 1)  '5.7
t_l.poi(2 * m(1)) = 0
Call search_for_two_line_value(t_l, m(0), n_(1), 1)
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
no% = two_line_value(k%).data(0).record.data1.index.i(m(0))
If no% > start% And two_line_value(no%).record_.no_reduce < 4 Then
'If is_two_record_related(two_line_value_, no%, two_line_value(no%).data(0).record, _
     line3_value_, t%, line3_value(t%).data(0).record) = False Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
'End If
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
combine_three_line_with_two_line = _
 combine_three_three_line_(two_line_value_, no%, _
  line3_value_, t%, m(0), n(0), 0)
If combine_three_line_with_two_line > 1 Then
 Exit Function
End If
Next k%
Next j%
Next i%
End Function

Public Function combine_line_value_with_two_line(ByVal lv%, _
         ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, k%, no%, last_tn%
Dim n_(1) As Integer
Dim m(1) As Integer
Dim tn() As Integer
Dim t_l As two_line_value_data0_type
Dim temp_record As total_record_type
temp_record.record_data.data0.condition_data.condition_no = 2
temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
temp_record.record_data.data0.condition_data.condition(2).ty = two_line_value_
temp_record.record_data.data0.condition_data.condition(1).no = lv%
temp_record.record_data.data0.theorem_no = 1
For i% = 0 To 1
m(0) = i%
m(1) = (i% + 1) Mod 2
t_l.poi(2 * m(0)) = line_value(lv%).data(0).data0.poi(0)
t_l.poi(2 * m(0) + 1) = line_value(lv%).data(0).data0.poi(1)
t_l.poi(2 * m(1)) = -1
Call search_for_two_line_value(t_l, m(0), n_(0), 1)  '5.7
t_l.poi(2 * m(1)) = 30000
Call search_for_two_line_value(t_l, m(0), n_(1), 1)
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
no% = two_line_value(k%).data(0).record.data1.index.i(m(0))
If no% > start% And two_line_value(no%).record_.no_reduce < 4 Then
 last_tn% = last_tn% + 1
 ReDim Preserve tn(last_tn%) As Integer
 tn(last_tn%) = no%
End If
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
temp_record.record_data.data0.condition_data.condition(2).no = no%
combine_line_value_with_two_line = set_line_value( _
   two_line_value(no%).data(0).data0.poi(2 * m(1)), two_line_value(no%).data(0).data0.poi(2 * m(1) + 1), _
      divide_string(minus_string(two_line_value(no%).data(0).data0.value, time_string( _
       line_value(lv%).data(0).data0.value, two_line_value(no%).data(0).data0.para(m(0)), False, False), False, False), _
        two_line_value(no%).data(0).data0.para(m(1)), True, False), two_line_value(no%).data(0).data0.n(2 * m(1)), _
         two_line_value(no%).data(0).data0.n(2 * m(1) + 1), two_line_value(no%).data(0).data0.line_no(m(1)), _
          temp_record, 0, no_reduce, False)
      Call set_level_(two_line_value(no%).record_.no_reduce, 4)
 If combine_line_value_with_two_line > 1 Then
  Exit Function
 End If
Next k%
Next i%
If InStr(1, line_value(lv%).data(0).data0.value_, "x", 0) > 0 Then
 For i% = 1 To last_conditions.last_cond(1).two_line_value_no
   combine_line_value_with_two_line = subs_line_value_to_two_line_value(i%, lv%)
    If combine_line_value_with_two_line > 1 Then
     Exit Function
    End If
 Next i%
End If
End Function

Public Function combine_line_value_with_three_line(ByVal lv%, _
    ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, k%, no%, last_tn%
Dim n_(1) As Integer
Dim m(2) As Integer
Dim tn() As Integer
Dim t_l As line3_value_data0_type
Dim temp_record As total_record_type
temp_record.record_data.data0.condition_data.condition_no = 2
temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
temp_record.record_data.data0.condition_data.condition(2).ty = line3_value_
temp_record.record_data.data0.condition_data.condition(1).no = lv%
temp_record.record_data.data0.theorem_no = 1
For i% = 0 To 2
m(0) = i%
m(1) = (i% + 1) Mod 3
m(2) = (i% + 2) Mod 3
t_l.poi(2 * m(0)) = line_value(lv%).data(0).data0.poi(0)
t_l.poi(2 * m(0) + 1) = line_value(lv%).data(0).data0.poi(1)
t_l.poi(2 * m(1)) = 0
Call search_for_line3_value(t_l, m(0), n_(0), 1)
t_l.poi(2 * m(1)) = 30000
Call search_for_line3_value(t_l, m(0), n_(1), 1)  '5.7
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
no% = line3_value(k%).data(0).record.data1.index.i(m(0))
If no% > start% And line3_value(no%).record_.no_reduce < 4 Then
 last_tn% = last_tn% + 1
 ReDim Preserve tn(last_tn%) As Integer
 tn(last_tn%) = no%
End If
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
 temp_record.record_data.data0.condition_data.condition(2).no = no%
  combine_line_value_with_three_line = set_two_line_value( _
    line3_value(no%).data(0).data0.poi(2 * m(1)), line3_value(no%).data(0).data0.poi(2 * m(1) + 1), _
     line3_value(no%).data(0).data0.poi(2 * m(2)), line3_value(no%).data(0).data0.poi(2 * m(2) + 1), _
    line3_value(no%).data(0).data0.n(2 * m(1)), line3_value(no%).data(0).data0.n(2 * m(1) + 1), _
     line3_value(no%).data(0).data0.n(2 * m(2)), line3_value(no%).data(0).data0.n(2 * m(2) + 1), _
    line3_value(no%).data(0).data0.line_no(m(1)), line3_value(no%).data(0).data0.line_no(m(2)), _
      line3_value(no%).data(0).data0.para(m(1)), line3_value(no%).data(0).data0.para(m(2)), _
       minus_string(line3_value(no%).data(0).data0.value, time_string( _
        line_value(lv%).data(0).data0.value, line3_value(no%).data(0).data0.para(m(0)), _
         False, False), True, False), temp_record, 0, no_reduce)
   Call set_level_(line3_value(no%).record_.no_reduce, 4)
  If combine_line_value_with_three_line > 1 Then
   Exit Function
  End If
 Next k%
Next i%
If InStr(1, line_value(lv%).data(0).data0.value_, "x", 0) > 0 Then
 For i% = 1 To last_conditions.last_cond(1).line3_value_no
   combine_line_value_with_three_line = subs_line_value_to_line3_value(i%, lv%)
    If combine_line_value_with_three_line > 1 Then
      Exit Function
    End If
 Next i%
End If
End Function
Public Function combine_relation_with_two_line(ByVal re%, _
            ByVal start%, no_reduce As Byte) As Byte
Dim i%, j%, k%, no%, last_tn%
Dim n(2) As Integer
Dim m(1) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim v(1) As String
Dim t_l As two_line_value_data0_type
If Drelation(re%).record_.no_reduce > 4 Then
 Exit Function
End If
For i% = 0 To 2
 n(0) = i%
  n(1) = (i% + 1) Mod 3
   n(2) = (i% + 2) Mod 3
 For j% = 0 To 1
  m(0) = j%
   m(1) = (j% + 1) Mod 2
 t_l.poi(2 * m(0)) = Drelation(re%).data(0).data0.poi(2 * n(0))
 t_l.poi(2 * m(0) + 1) = Drelation(re%).data(0).data0.poi(2 * n(0) + 1)
 t_l.poi(2 * m(1)) = -1
 Call search_for_two_line_value(t_l, m(0), n_(0), 1)
 t_l.poi(2 * m(1)) = 30000
 Call search_for_two_line_value(t_l, m(0), n_(1), 1)  '5.7
 last_tn% = 0
 For k% = n_(0) + 1 To n_(1)
 no% = two_line_value(k%).data(0).record.data1.index.i(m(0))
 If no% > start% And two_line_value(no%).record_.no_reduce < 4 Then
 'If is_two_record_related(two_line_value_, no%, two_line_value(no%).data(0).record, _
       relation_, re%, Drelation(re%).data(0).record) = False Then
 last_tn% = last_tn% + 1
  ReDim Preserve tn(last_tn%) As Integer
   tn(last_tn%) = no%
 End If
 'End If
 Next k%
 For k% = 1 To last_tn%
  no% = tn(k%)
  combine_relation_with_two_line = combine_relation_with_three_line_( _
    relation_, re%, two_line_value_, no%, n(0), m(0))
   If combine_relation_with_two_line > 1 Then
    Exit Function
   End If
 Next k%
Next j%
Next i%

End Function

Public Function combine_eline_with_two_line(ByVal el%, _
                  ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, no%, last_tn%
Dim n(1) As Integer
Dim m(1) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim t_l As two_line_value_data0_type
Dim temp_record As total_record_type
If Deline(el%).record_.no_reduce > 4 Then
 Exit Function
End If
temp_record.record_data.data0.condition_data.condition(2).ty = two_line_value_
  temp_record.record_data.data0.condition_data.condition(1).ty = eline_
   temp_record.record_data.data0.condition_data.condition(1).no = el%
       temp_record.record_data.data0.condition_data.condition_no = 2
        temp_record.record_data.data0.theorem_no = 1
For i% = 0 To 1
 n(0) = i%
  n(1) = (i% + 1) Mod 2
For j% = 0 To 1
   m(0) = j%
    m(1) = (j% + 1) Mod 2
t_l.poi(2 * m(0)) = Deline(el%).data(0).data0.poi(2 * n(0))
t_l.poi(2 * m(0) + 1) = Deline(el%).data(0).data0.poi(2 * n(0) + 1)
t_l.poi(2 * m(1)) = -1
Call search_for_two_line_value(t_l, m(0), n_(0), 1)  '5.7
t_l.poi(2 * m(1)) = 30000
Call search_for_two_line_value(t_l, m(0), n_(1), 1)
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
no% = two_line_value(k%).data(0).record.data1.index.i(m(0))
If no% > start% Then
'If is_two_record_related(two_line_value_, no%, two_line_value(no%).data(0).record, _
     eline_, el%, Deline(el%).data(0).record) = False And _
      two_line_value(no%).record_.no_reduce < 255 Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
'End If
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
temp_record.record_data.data0.condition_data.condition(2).no = no%
combine_eline_with_two_line = set_two_line_value( _
 Deline(el%).data(0).data0.poi(2 * n(1)), Deline(el%).data(0).data0.poi(2 * n(1) + 1), _
   two_line_value(no%).data(0).data0.poi(2 * m(1)), two_line_value(no%).data(0).data0.poi(2 * m(1) + 1), _
 Deline(el%).data(0).data0.n(2 * n(1)), Deline(el%).data(0).data0.n(2 * n(1) + 1), _
   two_line_value(no%).data(0).data0.n(2 * m(1)), two_line_value(no%).data(0).data0.n(2 * m(1) + 1), _
 Deline(el%).data(0).data0.line_no(n(1)), two_line_value(no%).data(0).data0.line_no(m(1)), _
    two_line_value(no%).data(0).data0.para(m(0)), two_line_value(no%).data(0).data0.para(m(1)), _
     two_line_value(no%).data(0).data0.value, temp_record, 0, no_reduce)
If combine_eline_with_two_line > 0 Then
Call set_level_(two_line_value(no%).record_.no_reduce, 4)
If combine_eline_with_two_line > 1 Then
  Exit Function
 End If
End If
Next k%
Next j%
Next i%

End Function

Public Function combine_line_value_with_eline(ByVal lv%, _
           ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, k%, no%, last_tn%
Dim n_(1) As Integer
Dim m(1) As Integer
Dim tn() As Integer
Dim e_l As eline_data0_type
Dim temp_record As total_record_type
temp_record.record_data.data0.condition_data.condition_no = 2
temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
temp_record.record_data.data0.condition_data.condition(2).ty = eline_
temp_record.record_data.data0.condition_data.condition(1).no = lv%
temp_record.record_data.data0.theorem_no = 1
For i% = 0 To 1
m(0) = i%
m(1) = (i% + 1) Mod 2
e_l.poi(2 * m(0)) = line_value(lv%).data(0).data0.poi(0)
e_l.poi(2 * m(0) + 1) = line_value(lv%).data(0).data0.poi(1)
e_l.poi(2 * m(1)) = -1
Call search_for_eline(e_l, m(0), n_(0), 1)
e_l.poi(2 * m(1)) = 30000
Call search_for_eline(e_l, m(0), n_(1), 1)  '5.7
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
no% = Deline(k%).data(0).record.data1.index.i(m(0))
If no% > start% And Deline(no%).record_.no_reduce < 4 Then
 last_tn% = last_tn% + 1
 ReDim Preserve tn(last_tn%) As Integer
 tn(last_tn%) = no%
End If
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
temp_record.record_data.data0.condition_data.condition(2).no = no%
combine_line_value_with_eline = _
 set_line_value(Deline(no%).data(0).data0.poi(2 * m(1)), _
  Deline(no%).data(0).data0.poi(2 * m(1) + 1), _
   line_value(lv%).data(0).data0.value, Deline(no%).data(0).data0.n(2 * m(1)), _
    Deline(no%).data(0).data0.n(2 * m(1) + 1), _
     Deline(no%).data(0).data0.line_no(m(1)), temp_record, 0, no_reduce, False)
If combine_line_value_with_eline > 0 Then
  Call set_level_(Deline(no%).record_.no_reduce, 4)
If combine_line_value_with_eline > 1 Then
 Exit Function
End If
End If
Next k%
Next i%
For i% = 1 To last_conditions.last_cond(1).eline_no
 If line_value(lv%).data(0).data0.line_no = Deline(i%).data(0).data0.line_no(0) Then
    combine_line_value_with_eline = combine_eline_with_line_value0(i%, lv%, 0)
     If combine_line_value_with_eline > 1 Then
      Exit Function
     End If
 ElseIf line_value(lv%).data(0).data0.line_no = Deline(i%).data(0).data0.line_no(1) Then
    combine_line_value_with_eline = combine_eline_with_line_value0(i%, lv%, 1)
     If combine_line_value_with_eline > 1 Then
      Exit Function
     End If
 End If
Next i%
End Function
Public Function combine_three_angle_with_three_angle(ByVal A3%, _
                             ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, l%, tn%
Dim ty As Byte
Dim n1_(1) As Integer
Dim n2_(1) As Integer
Dim A(1) As Integer
Dim temp_record As total_record_type
Dim t_A As angle3_value_data0_type
'If angle3_value(A3%).data(0).data0.reduce = True Then
temp_record.record_data.data0.condition_data.condition_no = 1
temp_record.record_data.data0.condition_data.condition(1).ty = angle3_value_
temp_record.record_data.data0.condition_data.condition(1).no = A3%
temp_record.record_data.data0.theorem_no = 1
t_A = angle3_value(A3%).data(0).data0
combine_three_angle_with_three_angle = _
   combine_three_angle_with_three_angle0(t_A, temp_record.record_data)
'End If
'For i% = 0 To 2
' If angle3_value(A3%).data(0).angle(i%) > 0 Then
'    A(0) = angle_number(Lin(angle(angle3_value(A3%).data(0).angle(i%)).data(0).line_no(0)). _
        data.poi((angle(angle3_value(A3%).data(0).angle(i%)).data(0).te(0) + 1) Mod 2), _
          angle(angle3_value(A3%).data(0).angle(i%)).data(0).poi(1), _
            angle(angle3_value(A3%).data(0).angle(i%)).data(0).poi(2), 0)
'    A(1) = angle_number(Lin(angle(angle3_value(A3%).data(0).angle(i%)).data(0).line_no(1)). _
        data.poi((angle(angle3_value(A3%).data(0).angle(i%)).data(0).te(1) + 1) Mod 2), _
          angle(angle3_value(A3%).data(0).angle(i%)).data(0).poi(1), _
            angle(angle3_value(A3%).data(0).angle(i%)).data(0).poi(0), 0)
'    If A(0) <> 0 Then
'     t_A = angle3_value(A3%).data
'      t_A.angle(i%) = Abs(A(0))
'       t_A.para(i%) = time_string("-1", t_A.para(i%))
'        t_A.value = add_string(t_A.value, time_string(t_A.para(i%), "180"))
'         combine_three_angle_with_three_angle = _
'          combine_three_angle_with_three_angle0(t_A, temp_record.record_data)
'          If combine_three_angle_with_three_angle > 1 Then
'           Exit Function
'          End If
'    End If
'    If A(1) <> 0 Then
'     t_A = angle3_value(A3%).data
'      t_A.angle(i%) = Abs(A(1))
'       t_A.para(i%) = time_string("-1", t_A.para(i%))
'        t_A.value = add_string(t_A.value, time_string(t_A.para(i%), "180"))
'         combine_three_angle_with_three_angle = _
          combine_three_angle_with_three_angle0(t_A, temp_record.record_data)
'           If combine_three_angle_with_three_angle > 1 Then
'            Exit Function
'           End If
'    End If
' End If
'Next i%
End Function

Public Function combine_point4_on_circle_with_point4_on_circle(ByVal c%, _
   ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, no%, t_n%, last_tn%
Dim n(3) As Integer
Dim m(3) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim temp_record As total_record_type
Dim p4_on_C As four_point_on_circle_data_type
If four_point_on_circle(c%).record_.no_reduce > 4 Then
 Exit Function
End If
temp_record.record_data.data0.condition_data.condition_no = 2
temp_record.record_data.data0.condition_data.condition(1).ty = point4_on_circle_
temp_record.record_data.data0.condition_data.condition(2).ty = point4_on_circle_
temp_record.record_data.data0.condition_data.condition(1).no = c%
temp_record.record_data.data0.theorem_no = 1
For i% = 0 To 5
If i% < 2 Then
n(0) = i%
 n(1) = (i% + 1) Mod 4
  n(2) = (i% + 2) Mod 4
   n(3) = (i% + 3) Mod 4
ElseIf i% = 4 Then
 n(0) = 0
  n(1) = 2
   n(2) = 3
    n(3) = 1
ElseIf i% = 5 Then
 n(0) = 0
  n(1) = 1
   n(2) = 3
    n(3) = 2
Else
GoTo combine_point4_on_circle_with_point4_on_circle_next1
End If
For j% = 0 To 6
If j% < 2 Then
m(0) = j%
 m(1) = (j% + 1) Mod 4
  m(2) = (j% + 2) Mod 4
   m(3) = (j% + 3) Mod 4
ElseIf j% = 4 Then
 m(0) = 0
  m(1) = 2
   m(2) = 3
    m(3) = 1
ElseIf j% = 5 Then
 m(0) = 0
  m(1) = 1
   m(2) = 3
    m(3) = 2
Else
GoTo combine_point4_on_circle_with_point4_on_circle_next2
End If
p4_on_C.poi(m(0)) = four_point_on_circle(c%).data(0).poi(n(0))
p4_on_C.poi(m(1)) = four_point_on_circle(c%).data(0).poi(n(1))
p4_on_C.poi(m(2)) = four_point_on_circle(c%).data(0).poi(n(2))
p4_on_C.poi(m(3)) = -1
Call search_for_four_point_on_circle(p4_on_C, 1, j%, n_(0), 1)
p4_on_C.poi(m(3)) = 30000
Call search_for_four_point_on_circle(p4_on_C, 1, j%, n_(1), 1)
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
no% = four_point_on_circle(k%).data(0).record.data1.index.i(j%)
If no% > 0 And no% <> c% And _
    four_point_on_circle(no%).record_.no_reduce < 4 Then
 last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
Next k%
For k% = 1 To last_tn%
 no% = tn(k%)
  temp_record.record_data.data0.condition_data.condition(2).no = no%
t_n% = 0
 combine_point4_on_circle_with_point4_on_circle = _
 set_four_point_on_circle(four_point_on_circle(c%).data(0).poi(n(1)), _
   four_point_on_circle(c%).data(0).poi(n(2)), four_point_on_circle(c%).data(0).poi(n(3)), _
     four_point_on_circle(no%).data(0).poi(n(3)), 0, temp_record, t_n%, no_reduce)
If combine_point4_on_circle_with_point4_on_circle > 0 Then
  Call set_level_(four_point_on_circle(t_n%).record_.no_reduce, 4)
If combine_point4_on_circle_with_point4_on_circle > 1 Then
 Exit Function
End If
End If
t_n% = 0
combine_point4_on_circle_with_point4_on_circle = _
 set_four_point_on_circle(four_point_on_circle(c%).data(0).poi(n(0)), _
   four_point_on_circle(c%).data(0).poi(n(2)), four_point_on_circle(c%).data(0).poi(n(3)), _
     four_point_on_circle(no%).data(0).poi(n(3)), 0, temp_record, t_n%, no_reduce)
If combine_point4_on_circle_with_point4_on_circle > 0 Then
  Call set_level_(four_point_on_circle(t_n%).record_.no_reduce, 4)
If combine_point4_on_circle_with_point4_on_circle > 1 Then
 Exit Function
End If
End If
t_n% = 0
combine_point4_on_circle_with_point4_on_circle = _
 set_four_point_on_circle(four_point_on_circle(c%).data(0).poi(n(0)), _
   four_point_on_circle(c%).data(0).poi(n(1)), four_point_on_circle(c%).data(0).poi(n(3)), _
     four_point_on_circle(no%).data(0).poi(n(3)), 0, temp_record, t_n%, no_reduce)
If combine_point4_on_circle_with_point4_on_circle > 0 Then
  Call set_level_(four_point_on_circle(t_n%).record_.no_reduce, 4)
If combine_point4_on_circle_with_point4_on_circle > 1 Then
 Exit Function
End If
End If
Next k%
combine_point4_on_circle_with_point4_on_circle_next2:
Next j%
combine_point4_on_circle_with_point4_on_circle_next1:
Next i%
End Function
Public Function combine_eline_with_relation(ByVal el%, _
               ByVal start%, ByVal no_reduce As Byte) As Byte '10.10
Dim i%, k%, l%, no%
Dim n(2) As Integer
Dim m(2) As Integer
Dim m0(2) As Integer
Dim n_(1) As Integer
Dim s(1) As String
Dim t(1) As String
Dim tn() As Integer
Dim re As relation_data0_type
Dim last_tn As Integer
If Deline(el%).record_.no_reduce > 4 Then
 Exit Function
End If
For k% = 0 To 1
n(0) = k%
n(1) = (k% + 1) Mod 2
For l% = 0 To 2
m(0) = l%
m(1) = (l% + 1) Mod 3
m(2) = (l% + 2) Mod 3
re.poi(2 * m(0)) = Deline(el%).data(0).data0.poi(2 * n(0))
re.poi(2 * m(0) + 1) = Deline(el%).data(0).data0.poi(2 * n(0) + 1)
re.poi(2 * m(1)) = -1
Call search_for_relation(re, m(0), n_(0), 1)  '5.7
re.poi(2 * m(1)) = 30000
Call search_for_relation(re, m(0), n_(1), 1)
last_tn = 0
For i% = n_(0) + 1 To n_(1)
no% = Drelation(i%).data(0).record.data1.index.i(m(0))
If no% > start% And Drelation(no%).record_.no_reduce < 4 Then
'If is_two_record_related(relation_, no%, Drelation(no%).data(0).record, _
        eline_, el%, Deline(el%).data(0).record) = False Then
last_tn = last_tn + 1
ReDim Preserve tn(last_tn) As Integer
tn(last_tn) = no%
End If
'End If
Next i%
For i% = 1 To last_tn
no% = tn(i%)
combine_eline_with_relation = _
  combine_relation_with_relation_(eline_, el%, relation_, no%, n(0), m(0))
If combine_eline_with_relation > 1 Then
 Exit Function
End If
Next i%
Next l%
Next k%
End Function

Public Function combine_mid_point_with_mid_point0(ByVal md%, _
      k%, l%, ByVal no_reduce) As Byte
Dim i%, no%
Dim n(5) As Integer
Dim m(5) As Integer
Dim s(1) As String
Dim t(1) As String
Dim md_p  As mid_point_data0_type
Dim temp_record As total_record_type
Dim re As total_record_type
If Dmid_point(md%).record_.no_reduce > 4 Then
 Exit Function
End If
Call add_conditions_to_record(midpoint_, md%, 0, 0, re.record_data.data0.condition_data)
re.record_data.data0.theorem_no = 1
Call read_ratio_from_relation("1", k%, s(0), s(1), True, 3)
If k% = 0 Then
n(0) = 0
n(1) = 1
n(2) = 1
n(3) = 2
n(4) = 0
n(5) = 2
ElseIf k% = 1 Then
n(0) = 1
n(1) = 2
n(2) = 0
n(3) = 2
n(4) = 0
n(5) = 1
ElseIf k% = 2 Then
n(0) = 0
n(1) = 2
n(2) = 0
n(3) = 1
n(4) = 1
n(5) = 2
End If
If l% = 0 Then
m(0) = 0
m(1) = 1
m(2) = 1
m(3) = 2
m(4) = 0
m(5) = 2
ElseIf l% = 1 Then
m(0) = 1
m(1) = 2
m(2) = 0
m(3) = 2
m(4) = 0
m(5) = 1
Else
m(0) = 0
m(1) = 2
m(2) = 0
m(3) = 1
m(4) = 1
m(5) = 2
End If
md_p.poi(m(0)) = Dmid_point(md%).data(0).data0.poi(n(0))
md_p.poi(m(1)) = Dmid_point(md%).data(0).data0.poi(n(1))
If search_for_mid_point(md_p, l%, no%, 2) Then  '5.7原l%+3
If no% > 0 And no% < md% And _
  Dmid_point(no%).record_.no_reduce < 4 Then
  ' If is_two_record_related(midpoint_, no%, Dmid_point(no%).data(0).record, _
       midpoint_, md%, Dmid_point(md%).data(0).record) = False Then
Call read_ratio_from_relation("1", l%, t(0), t(1), True, 3)
temp_record = re
Call add_conditions_to_record(midpoint_, no%, 0, 0, temp_record.record_data.data0.condition_data)
'******
combine_mid_point_with_mid_point0 = set_Drelation( _
 Dmid_point(no%).data(0).data0.poi(m(2)), Dmid_point(no%).data(0).data0.poi(m(3)), _
  Dmid_point(md%).data(0).data0.poi(n(2)), Dmid_point(md%).data(0).data0.poi(n(3)), _
    Dmid_point(no%).data(0).data0.n(m(2)), Dmid_point(no%).data(0).data0.n(m(3)), _
     Dmid_point(md%).data(0).data0.n(n(2)), Dmid_point(md%).data(0).data0.n(n(3)), _
      Dmid_point(no%).data(0).data0.line_no, Dmid_point(md%).data(0).data0.line_no, _
       divide_string(s(0), t(0), True, False), _
        temp_record, 0, 0, 0, 0, no_reduce, False)
If combine_mid_point_with_mid_point0 > 1 Then
 Exit Function
End If
combine_mid_point_with_mid_point0 = _
set_Drelation(Dmid_point(no%).data(0).data0.poi(m(2)), _
 Dmid_point(no%).data(0).data0.poi(m(3)), Dmid_point(md%).data(0).data0.poi(n(4)), _
  Dmid_point(md%).data(0).data0.poi(n(5)), Dmid_point(no%).data(0).data0.n(m(2)), _
   Dmid_point(no%).data(0).data0.n(m(3)), Dmid_point(md%).data(0).data0.n(n(4)), _
    Dmid_point(md%).data(0).data0.n(n(5)), Dmid_point(no%).data(0).data0.line_no, _
     Dmid_point(md%).data(0).data0.line_no, divide_string(s(1), t(0), True, False), _
      temp_record, 0, 0, 0, 0, no_reduce, False)
If combine_mid_point_with_mid_point0 > 1 Then
 Exit Function
End If
combine_mid_point_with_mid_point0 = _
 set_Drelation( _
  Dmid_point(no%).data(0).data0.poi(m(4)), Dmid_point(no%).data(0).data0.poi(m(5)), _
   Dmid_point(md%).data(0).data0.poi(n(2)), Dmid_point(md%).data(0).data0.poi(n(3)), _
    Dmid_point(no%).data(0).data0.n(m(4)), Dmid_point(no%).data(0).data0.n(m(5)), _
     Dmid_point(md%).data(0).data0.n(n(2)), Dmid_point(md%).data(0).data0.n(n(3)), _
      Dmid_point(no%).data(0).data0.line_no, Dmid_point(md%).data(0).data0.line_no, _
       divide_string(s(0), t(1), True, False), _
        temp_record, 0, 0, 0, 0, no_reduce, False)
If combine_mid_point_with_mid_point0 > 1 Then
 Exit Function
End If
combine_mid_point_with_mid_point0 = _
 set_Drelation( _
  Dmid_point(no%).data(0).data0.poi(m(4)), Dmid_point(no%).data(0).data0.poi(m(5)), _
   Dmid_point(md%).data(0).data0.poi(n(4)), Dmid_point(md%).data(0).data0.poi(n(5)), _
    Dmid_point(no%).data(0).data0.n(m(4)), Dmid_point(no%).data(0).data0.n(m(5)), _
     Dmid_point(md%).data(0).data0.n(n(4)), Dmid_point(md%).data(0).data0.n(n(5)), _
      Dmid_point(no%).data(0).data0.line_no, Dmid_point(md%).data(0).data0.line_no, _
       divide_string(s(1), t(1), True, False), _
        temp_record, 0, 0, 0, 0, no_reduce, False)
If combine_mid_point_with_mid_point0 > 1 Then
 Exit Function
End If
End If
End If
'End If
End Function

Public Function combine_relation_with_relation_(ByVal ty1 As Byte, ByVal re1%, _
    ByVal ty2 As Byte, ByVal re2%, k%, l%) As Byte
Dim i%, no%
Dim ty As Byte
Dim p_(5) As Integer
Dim n_(5) As Integer
Dim p1(5) As Integer
Dim n1(5) As Integer
Dim l1(2) As Integer
Dim p2(5) As Integer
Dim n2(5) As Integer
Dim l2(2) As Integer
Dim v1(1) As String
Dim v2(1) As String
Dim v(3) As String
Dim temp_record As total_record_type
Dim is_no_initial As Integer
Dim c_data As condition_data_type
Call add_conditions_to_record(ty1, re1%, 0, 0, temp_record.record_data.data0.condition_data)
Call add_conditions_to_record(ty2, re2%, 0, 0, temp_record.record_data.data0.condition_data)
temp_record.record_data.data0.theorem_no = 1
'读出对应指标的比和点
 Call read_point_and_ratio_from_relation(ty1, re1%, k%, p1(), n1(), _
       l1(), v1(0), v1(1))
 Call read_point_and_ratio_from_relation(ty2, re2%, l%, p2(), n2(), _
       l2(), v2(0), v2(1))
If ty1 = line_value_ And ty2 = line_value_ Then
 If arrange_four_point(p1(0), p1(1), p2(0), p2(1), n1(0), n1(1), _
     n2(0), n2(1), l1(0), l2(0), p1(2), p1(3), p1(4), p1(5), 0, _
       0, n1(2), n1(3), n1(4), n1(5), 0, 0, l1(1), l1(2), 0, _
        ty, c_data, is_no_initial) Then
  If ty = 3 Or ty = 5 Then
   If is_no_initial = 1 Then
    Call add_record_to_record(c_data, temp_record.record_data.data0.condition_data)
   End If
   combine_relation_with_relation_ = set_line_value( _
    p1(2), p1(5), add_string(v1(0), v2(0), True, False), n1(2), n1(5), _
         l1(1), temp_record, 0, 0, False)
     If combine_relation_with_relation_ > 1 Then
      Exit Function
     End If
  ElseIf ty = 4 Then
   combine_relation_with_relation_ = set_line_value( _
    p1(2), p1(3), minus_string(v1(0), v2(0), True, False), n1(2), n2(3), _
      l1(1), temp_record, 0, 0, False)
     If combine_relation_with_relation_ > 1 Then
      Exit Function
     End If
  ElseIf ty = 6 Then
   combine_relation_with_relation_ = set_line_value( _
    p1(2), p1(3), minus_string(v2(0), v1(0), True, False), n1(2), n1(3), _
       l1(1), temp_record, 0, 0, False)
     If combine_relation_with_relation_ > 1 Then
      Exit Function
     End If
  ElseIf ty = 7 Then
   combine_relation_with_relation_ = set_line_value( _
    p1(4), p1(5), minus_string(v2(0), v1(0), True, False), n1(4), n1(5), _
      l1(2), temp_record, 0, 0, False)
     If combine_relation_with_relation_ > 1 Then
      Exit Function
     End If
  ElseIf ty = 8 Then
   combine_relation_with_relation_ = set_line_value( _
    p1(4), p1(5), minus_string(v1(0), v2(0), True, False), n1(4), n1(5), _
      l1(2), temp_record, 0, 0, False)
     If combine_relation_with_relation_ > 1 Then
      Exit Function
     End If
  End If
 End If
ElseIf ty1 = line_value_ Then
 If v2(0) <> "" Then
  combine_relation_with_relation_ = set_line_value( _
   p2(2), p2(3), divide_string(v1(0), v2(0), True, False), n2(2), n2(3), _
    l2(1), temp_record, 0, 0, False)
  If combine_relation_with_relation_ > 1 Then
   Exit Function
  End If
 End If
 If v2(1) <> "" Then
  combine_relation_with_relation_ = set_line_value( _
   p2(4), p2(5), divide_string(v1(0), v2(1), True, False), n2(4), n2(5), _
      l2(2), temp_record, 0, 0, False)
  If combine_relation_with_relation_ > 1 Then
   Exit Function
  End If
 End If
ElseIf ty2 = line_value_ Then
 If v1(0) <> "" Then
  combine_relation_with_relation_ = set_line_value( _
   p1(2), p1(3), divide_string(v2(0), v1(0), True, False), n1(2), n1(3), _
      l1(1), temp_record, 0, 0, False)
  If combine_relation_with_relation_ > 1 Then
   Exit Function
  End If
 End If
 If v1(1) <> "" Then
  combine_relation_with_relation_ = set_line_value( _
   p1(4), p1(5), divide_string(v2(0), v1(1), True, False), n1(4), n1(5), _
     l1(2), temp_record, 0, 0, False)
  If combine_relation_with_relation_ > 1 Then
   Exit Function
  End If
 End If
Else
If v1(0) <> "" And v2(0) <> "" Then
 combine_relation_with_relation_ = set_Drelation(p1(2), _
   p1(3), p2(2), p2(3), n1(2), n1(3), n2(2), n2(3), l1(1), l2(1), _
    divide_string(v2(0), v1(0), True, False), temp_record, 0, 0, 0, 0, 0, False)
 If combine_relation_with_relation_ > 1 Then
  Exit Function
 End If
End If
If v1(0) <> "" And v2(1) <> "" Then
 combine_relation_with_relation_ = set_Drelation(p1(2), _
   p1(3), p2(4), p2(5), n1(2), n1(3), n2(4), n2(5), l1(1), l2(2), _
    divide_string(v2(1), v1(0), True, False), temp_record, 0, 0, 0, 0, 0, False)
 If combine_relation_with_relation_ > 1 Then
  Exit Function
 End If
End If
If v1(1) <> "" And v2(0) <> "" Then
 combine_relation_with_relation_ = set_Drelation(p1(4), _
   p1(5), p2(2), p2(3), n1(4), n1(5), n2(2), n2(3), l1(2), l2(1), _
    divide_string(v2(0), v1(1), True, False), temp_record, 0, 0, 0, 0, 0, False)
 If combine_relation_with_relation_ > 1 Then
  Exit Function
 End If
End If
If v1(1) <> "" And v2(1) <> "" Then
 combine_relation_with_relation_ = set_Drelation(p1(4), _
   p1(5), p2(4), p2(5), n1(4), n1(5), n2(4), n2(5), l1(2), l2(2), _
    divide_string(v2(1), v1(1), True, False), temp_record, 0, 0, 0, 0, 0, False)
 If combine_relation_with_relation_ > 1 Then
  Exit Function
 End If
End If
'共线比例线段合并
  If ty1 = relation_ Then
     p1(0) = Drelation(re1%).data(0).data0.poi(0)
     p1(1) = Drelation(re1%).data(0).data0.poi(1)
     p1(2) = Drelation(re1%).data(0).data0.poi(2)
     p1(3) = Drelation(re1%).data(0).data0.poi(3)
     p1(4) = Drelation(re1%).data(0).data0.poi(4)
     p1(5) = Drelation(re1%).data(0).data0.poi(5)
     n1(0) = Drelation(re1%).data(0).data0.n(0)
     n1(1) = Drelation(re1%).data(0).data0.n(1)
     n1(2) = Drelation(re1%).data(0).data0.n(2)
     n1(3) = Drelation(re1%).data(0).data0.n(3)
     n1(4) = Drelation(re1%).data(0).data0.n(4)
     n1(5) = Drelation(re1%).data(0).data0.n(5)
     v1(0) = Drelation(re1%).data(0).data0.value
  ElseIf ty1 = midpoint_ Then
     p1(0) = Dmid_point(re1%).data(0).data0.poi(0)
     p1(1) = Dmid_point(re1%).data(0).data0.poi(1)
     p1(2) = Dmid_point(re1%).data(0).data0.poi(1)
     p1(3) = Dmid_point(re1%).data(0).data0.poi(2)
     p1(4) = Dmid_point(re1%).data(0).data0.poi(0)
     p1(5) = Dmid_point(re1%).data(0).data0.poi(2)
     n1(0) = Dmid_point(re1%).data(0).data0.n(0)
     n1(1) = Dmid_point(re1%).data(0).data0.n(1)
     n1(2) = Dmid_point(re1%).data(0).data0.n(1)
     n1(3) = Dmid_point(re1%).data(0).data0.n(2)
     n1(4) = Dmid_point(re1%).data(0).data0.n(0)
     p1(5) = Dmid_point(re1%).data(0).data0.n(2)
      v1(0) = "1"
 End If
  If ty2 = relation_ Then
     p2(0) = Drelation(re2%).data(0).data0.poi(0)
     p2(1) = Drelation(re2%).data(0).data0.poi(1)
     p2(2) = Drelation(re2%).data(0).data0.poi(2)
     p2(3) = Drelation(re2%).data(0).data0.poi(3)
     p2(4) = Drelation(re2%).data(0).data0.poi(4)
     p2(5) = Drelation(re2%).data(0).data0.poi(5)
     n2(0) = Drelation(re2%).data(0).data0.n(0)
     n2(1) = Drelation(re2%).data(0).data0.n(1)
     n2(2) = Drelation(re2%).data(0).data0.n(2)
     n2(3) = Drelation(re2%).data(0).data0.n(3)
     n2(4) = Drelation(re2%).data(0).data0.n(4)
     n2(5) = Drelation(re2%).data(0).data0.n(5)
     v2(0) = Drelation(re2%).data(0).data0.value
  ElseIf ty2 = midpoint_ Then
     p2(0) = Dmid_point(re2%).data(0).data0.poi(0)
     p2(1) = Dmid_point(re2%).data(0).data0.poi(1)
     p2(2) = Dmid_point(re2%).data(0).data0.poi(1)
     p2(3) = Dmid_point(re2%).data(0).data0.poi(2)
     p2(4) = Dmid_point(re2%).data(0).data0.poi(0)
     p2(5) = Dmid_point(re2%).data(0).data0.poi(2)
     n2(0) = Dmid_point(re2%).data(0).data0.n(0)
     n2(1) = Dmid_point(re2%).data(0).data0.n(1)
     n2(2) = Dmid_point(re2%).data(0).data0.n(1)
     n2(3) = Dmid_point(re2%).data(0).data0.n(2)
     n2(4) = Dmid_point(re2%).data(0).data0.n(0)
     n2(5) = Dmid_point(re2%).data(0).data0.n(2)
     v2(0) = "1"
  End If
If p1(4) > 0 And p1(5) > 0 And p2(4) > 0 And p2(5) > 0 Then
   If l1(0) = l2(0) Then '共线比例线段合并
     temp_record.record_data.data0.condition_data.condition_no = 2
     temp_record.record_data.data0.condition_data.condition(1).ty = ty1
     temp_record.record_data.data0.condition_data.condition(2).ty = ty2
     temp_record.record_data.data0.condition_data.condition(1).no = re1%
     temp_record.record_data.data0.condition_data.condition(2).no = re2%
     temp_record.record_data.data0.theorem_no = 1
     If k% = 0 And l% = 0 Then
        If n1(3) < n2(3) Then
         v(0) = time_string(v1(0), v2(0), True, False)
         v(1) = v2(0)
         v(2) = minus_string(v1(0), v2(0), True, False)
         v(3) = add_string(v(0), v(1), False, False)
         v(3) = add_string(v(3), v(2), True, False)
         combine_relation_with_relation_ = set_Drelation(p1(0), p1(3), p1(3), p2(3), _
            n1(0), n1(3), n1(3), n2(3), l1(0), l1(0), divide_string( _
             add_string(v(0), v(1), False, False), v(2), True, False), _
              temp_record, 0, 0, 0, 0, 0, False)
           If combine_relation_with_relation_ > 1 Then
             Exit Function
           End If
         combine_relation_with_relation_ = set_Drelation(p1(1), p1(3), p1(0), p2(3), _
            n1(1), n1(3), n1(0), n2(3), l1(0), l1(0), divide_string( _
               v(1), v(3), True, False), _
              temp_record, 0, 0, 0, 0, 0, False)
           If combine_relation_with_relation_ > 1 Then
             Exit Function
           End If
        Else
         v(0) = time_string(v1(0), v2(0), True, False)
         v(1) = v1(0)
         v(2) = minus_string(v2(0), v1(0), True, False)
         v(3) = add_string(v(0), v(1), False, False)
         v(3) = add_string(v(3), v(2), True, False)
        combine_relation_with_relation_ = set_Drelation(p2(0), p2(3), p2(3), p1(3), _
           n2(0), n2(3), n2(3), n1(3), l1(0), l1(0), divide_string( _
             add_string(v(0), v(1), False, False), v(2), True, False), _
              temp_record, 0, 0, 0, 0, 0, False)
           If combine_relation_with_relation_ > 1 Then
              Exit Function
           End If
        combine_relation_with_relation_ = set_Drelation(p2(1), p2(3), p2(0), p1(3), _
            n2(1), n2(3), n2(0), n1(3), l1(0), l1(0), divide_string( _
               v(1), v(3), True, False), _
              temp_record, 0, 0, 0, 0, 0, False)
           If combine_relation_with_relation_ > 1 Then
             Exit Function
           End If
       End If
     ElseIf k% = 1 And l% = 1 Then
        If n1(0) < n2(0) Then
         v(0) = minus_string(v1(0), v2(0), True, False)
         v(1) = v2(0)
         v(2) = "1"
         v(3) = add_string(v(0), v(1), False, False)
         v(3) = add_string(v(3), v(2), True, False)
         combine_relation_with_relation_ = set_Drelation(p1(0), p2(0), p2(0), p2(3), _
             n1(0), n2(0), n2(0), n2(3), l1(0), l1(0), divide_string( _
               v(0), add_string(v(1), v(2), False, False), True, False), _
                 temp_record, 0, 0, 0, 0, 0, False)
           If combine_relation_with_relation_ > 1 Then
             Exit Function
           End If
         combine_relation_with_relation_ = set_Drelation(p2(0), p2(1), p1(0), p2(3), _
            n2(0), n2(1), n1(0), n2(3), l1(0), l1(0), divide_string( _
               v(1), v(3), True, False), _
              temp_record, 0, 0, 0, 0, 0, False)
           If combine_relation_with_relation_ > 1 Then
             Exit Function
           End If
       Else
         v(0) = minus_string(v2(0), v1(0), True, False)
         v(1) = v1(0)
         v(2) = "1"
         v(3) = add_string(v(0), v(1), False, False)
         v(3) = add_string(v(3), v(2), True, False)
         combine_relation_with_relation_ = set_Drelation(p2(0), p1(0), p1(0), p1(3), _
            n2(0), n1(0), n1(0), n1(3), l1(0), l1(0), divide_string( _
              v(0), add_string(v(1), v(2), False, False), True, False), _
               temp_record, 0, 0, 0, 0, 0, False)
           If combine_relation_with_relation_ > 1 Then
             Exit Function
           End If
         combine_relation_with_relation_ = set_Drelation(p1(0), p1(1), p2(0), p1(3), _
            n1(0), n1(1), n2(0), n1(3), l1(0), l1(0), divide_string( _
               v(1), v(3), True, False), _
              temp_record, 0, 0, 0, 0, 0, False)
           If combine_relation_with_relation_ > 1 Then
             Exit Function
           End If
        End If
     ElseIf k% = 2 And l% = 2 Then
      If n1(1) < n2(1) Then
        v(2) = add_string("1", v2(0), True, False)
        v(0) = time_string(v1(0), v(2), True, False)
        v(1) = minus_string(v2(0), v1(0), True, False)
        v(2) = add_string("1", v1(0), True, False)
        v(0) = add_string(v(0), v(1), False, False)
        v(2) = add_string(v(0), v(2), False, False)
      combine_relation_with_relation_ = set_Drelation(p1(1), p2(1), p1(0), p1(3), _
         n1(1), n2(1), n1(0), n1(3), l1(0), l1(0), divide_string(v(1), v(2), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
         If combine_relation_with_relation_ > 1 Then
           Exit Function
         End If
        Else
        v(2) = add_string("1", v1(0), True, False)
        v(0) = time_string(v2(0), v(2), True, False)
        v(1) = minus_string(v1(0), v2(0), True, False)
        v(2) = add_string("1", v2(0), True, False)
        v(0) = add_string(v(0), v(1), False, False)
        v(2) = add_string(v(0), v(2), False, False)
      combine_relation_with_relation_ = set_Drelation(p2(1), p1(1), p1(0), p1(3), _
         n2(1), n1(1), n1(0), n1(3), l1(0), l1(0), divide_string(v(1), v(2), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
         If combine_relation_with_relation_ > 1 Then
           Exit Function
         End If
        End If
     ElseIf k% = 0 And l% = 1 Then
      v(0) = time_string(v2(0), v1(0), True, False)
      v(1) = v1(0)
      v(2) = "1"
      combine_relation_with_relation_ = set_Drelation(p2(0), p1(1), p1(1), p1(3), _
         n2(0), n1(1), n1(1), n1(3), l1(0), l1(0), divide_string(add_string(v(0), v(1), False, False), v(2), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
      If combine_relation_with_relation_ > 1 Then
         Exit Function
      End If
      combine_relation_with_relation_ = set_Drelation(p2(0), p2(1), p2(1), p1(3), _
         n2(0), n2(1), n2(1), n1(3), l1(0), l1(0), divide_string(v(0), add_string(v(1), v(2), False, False), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
      If combine_relation_with_relation_ > 1 Then
         Exit Function
      End If
         v(3) = add_string(v(0), v(1), False, False)
         v(3) = add_string(v(3), v(2), True, False)
      combine_relation_with_relation_ = set_Drelation(p1(0), p1(1), p2(0), p1(3), _
         n1(0), n1(1), n2(0), n1(3), l1(0), l1(0), divide_string(v(1), v(3), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
      If combine_relation_with_relation_ > 1 Then
         Exit Function
      End If
    ElseIf k% = 1 And l% = 0 Then
      v(0) = time_string(v1(0), v2(0), True, False)
      v(1) = v2(0)
      v(2) = "1"
      combine_relation_with_relation_ = set_Drelation(p1(0), p2(1), p2(1), p2(3), _
         n1(0), n2(1), n2(1), n2(3), l1(0), l1(0), divide_string(add_string(v(0), v(1), False, False), v(2), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
      If combine_relation_with_relation_ > 1 Then
         Exit Function
      End If
      combine_relation_with_relation_ = set_Drelation(p1(0), p1(1), p1(1), p2(3), _
         n1(0), n1(1), n1(1), n2(3), l1(0), l1(0), divide_string(v(0), add_string(v(1), v(2), False, False), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
      If combine_relation_with_relation_ > 1 Then
         Exit Function
      End If
         v(3) = add_string(v(0), v(1), False, False)
         v(3) = add_string(v(3), v(2), True, False)
      combine_relation_with_relation_ = set_Drelation(p2(0), p2(1), p1(0), p2(3), _
         n2(0), n2(1), n1(0), n2(3), l1(0), l1(0), divide_string(v(1), v(3), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
      If combine_relation_with_relation_ > 1 Then
         Exit Function
      End If
  ElseIf k% = 0 And l% = 2 Then
      v(2) = add_string("1", v2(0), True, False)
      v(0) = time_string(v2(0), v1(0), True, False)
      v(1) = v1(0)
      combine_relation_with_relation_ = set_Drelation(p2(0), p2(1), p2(1), p1(3), _
         n2(0), n2(1), n2(1), n1(3), l1(0), l1(0), divide_string(v(0), add_string(v(1), v(2), False, False), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
       If combine_relation_with_relation_ > 1 Then
         Exit Function
       End If
         v(3) = add_string(v(0), v(1), False, False)
         v(3) = add_string(v(3), v(2), True, False)
      combine_relation_with_relation_ = set_Drelation(p2(2), p2(3), p1(0), p1(3), _
         n2(2), n2(3), n1(0), n1(3), l1(0), l1(0), divide_string(v(1), v(3), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
      If combine_relation_with_relation_ > 1 Then
         Exit Function
      End If
   ElseIf k% = 2 And l% = 0 Then
      v(2) = add_string("1", v1(0), True, False)
      v(0) = time_string(v1(0), v2(0), True, False)
      v(1) = v2(0)
      combine_relation_with_relation_ = set_Drelation(p1(0), p1(1), p1(1), p2(3), _
         n1(0), n1(1), n1(1), n2(3), l1(0), l1(0), divide_string(v(0), add_string(v(1), v(2), False, False), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
       If combine_relation_with_relation_ > 1 Then
         Exit Function
       End If
         v(3) = add_string(v(0), v(1), False, False)
         v(3) = add_string(v(3), v(2), True, False)
      combine_relation_with_relation_ = set_Drelation(p1(2), p1(3), p2(0), p2(3), _
         n1(2), n1(3), n2(0), n2(3), l1(0), l1(0), divide_string(v(1), v(3), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
      If combine_relation_with_relation_ > 1 Then
         Exit Function
      End If
    ElseIf k% = l And l% = 2 Then
      v(0) = add_string("1", v2(0), False, False)
      v(0) = time_string(v1(0), v(0), True, False)
      v(1) = v2(0)
      v(2) = "1"
      combine_relation_with_relation_ = set_Drelation(p1(0), p2(1), p2(1), p2(3), _
        n1(0), n2(1), n2(1), n2(3), l1(0), l1(0), add_string(v(0), v(1), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
       If combine_relation_with_relation_ > 1 Then
         Exit Function
       End If
         v(3) = add_string(v(0), v(1), False, False)
         v(3) = add_string(v(3), v(2), True, False)
      combine_relation_with_relation_ = set_Drelation(p2(0), p2(1), p1(0), p1(3), _
         n2(0), n2(1), n1(0), n1(3), l1(0), l1(0), divide_string(v(1), v(3), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
      If combine_relation_with_relation_ > 1 Then
         Exit Function
      End If
    ElseIf k% = 2 And l% = 1 Then
      v(0) = add_string("1", v1(0), False, False)
      v(0) = time_string(v2(0), v(0), True, False)
      v(1) = v1(0)
      v(2) = "1"
      combine_relation_with_relation_ = set_Drelation(p2(0), p1(1), p1(1), p1(3), _
        n2(0), n1(1), n1(1), n1(3), l1(0), l1(0), add_string(v(0), v(1), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
       If combine_relation_with_relation_ > 1 Then
         Exit Function
       End If
         v(3) = add_string(v(0), v(1), False, False)
         v(3) = add_string(v(3), v(2), True, False)
      combine_relation_with_relation_ = set_Drelation(p1(0), p1(1), p2(0), p2(3), _
         n1(0), n1(1), n2(0), n2(3), l1(0), l1(0), divide_string(v(1), v(3), True, False), _
           temp_record, 0, 0, 0, 0, 0, False)
      If combine_relation_with_relation_ > 1 Then
         Exit Function
      End If
     End If
   End If
End If
'If k% = 2 And l% = 2 And l1(1) = l2(1) Then
 'If ty1 = relation_ Then
 'v1(0) = add_string(Drelation(re1%).data(0).data0.value, "1", False, False)
 'v1(1) = divide_string("1", v1(0), True, False)
 'v1(0) = divide_string(Drelation(re1%).data(0).data0.value, v1(0), True, False)
 'p1(0) = Drelation(re1%).data(0).data0.poi(0)
 'p1(1) = Drelation(re1%).data(0).data0.poi(1)
 'p1(2) = Drelation(re1%).data(0).data0.poi(3)
 'n1(0) = Drelation(re1%).data(0).data0.n(0)
 'n1(1) = Drelation(re1%).data(0).data0.n(1)
 'n1(2) = Drelation(re1%).data(0).data0.n(3)
 'Else
 'v1(0) = "1/2" 'add_string(Drelation(re1%).data(0).data0.value, "1", False, False)
 'v1(1) = "1/2" 'divide_string("1", v1(0), True, False)
 'p1(0) = Dmid_point(re1%).data(0).data0.poi(0)
 'p1(1) = Dmid_point(re1%).data(0).data0.poi(1)
 'p1(2) = Dmid_point(re1%).data(0).data0.poi(2)
 'n1(0) = Dmid_point(re1%).data(0).data0.n(0)
 'n1(1) = Dmid_point(re1%).data(0).data0.n(1)
 'n1(2) = Dmid_point(re1%).data(0).data0.n(2)
 'End If
 'If ty2 = relation_ Then
 'v2(0) = add_string(Drelation(re2%).data(0).data0.value, "1", False, False)
 'v2(1) = divide_string("1", v2(0), True, False)
 'v2(0) = divide_string(Drelation(re2%).data(0).data0.value, v2(0), True, False)
 'p2(0) = Drelation(re2%).data(0).data0.poi(0)
 'p2(1) = Drelation(re2%).data(0).data0.poi(1)
 'p2(2) = Drelation(re2%).data(0).data0.poi(3)
 'n2(0) = Drelation(re2%).data(0).data0.n(0)
 'n2(1) = Drelation(re2%).data(0).data0.n(1)
 'n2(2) = Drelation(re2%).data(0).data0.n(3)
 'Else
 'v2(0) = "1/2" 'add_string(Drelation(re1%).data(0).data0.value, "1", False, False)
 'v2(1) = "1/2" 'divide_string("1", v1(0), True, False)
 'p2(0) = Dmid_point(re2%).data(0).data0.poi(0)
 'p2(1) = Dmid_point(re2%).data(0).data0.poi(1)
 'p2(2) = Dmid_point(re2%).data(0).data0.poi(2)
 'n2(0) = Dmid_point(re2%).data(0).data0.n(0)
 'n2(1) = Dmid_point(re2%).data(0).data0.n(1)
 'n2(2) = Dmid_point(re2%).data(0).data0.n(2)
 'End If
 'p_(0) = p1(0)
 'p_(3) = p1(2)
 'n_(0) = n1(0)
 'n_(3) = n1(2)
 'If n1(1) < n2(1) Then
 'p_(1) = p1(1)
 'p_(2) = p2(1)
 'n_(1) = n1(1)
 'n_(2) = n2(1)
 'v(0) = v1(0)
 'v(1) = minus_string(v2(0), v1(0), True, False)
 'v(2) = v2(1)
 'Else
 'p_(1) = p2(1)
 'p_(2) = p1(1)
 'n_(1) = n2(1)
 'n_(2) = n1(1)
 'v(0) = v2(0)
 'v(1) = minus_string(v1(0), v2(0), True, False)
 'v(2) = v1(1)
 'End If
 'combine_relation_with_relation_ = set_Drelation(p_(0), _
   p_(1), p_(1), p_(2), n_(0), n_(1), n_(1), n_(2), l1(2), l2(1), _
    divide_string(v(0), v(1), True, False), temp_record, 0, 0, 0, 0, 0)
 'If combine_relation_with_relation_ > 1 Then
 ' Exit Function
 'End If
 'combine_relation_with_relation_ = set_Drelation(p_(1), _
   p_(2), p_(2), p_(3), n_(1), n_(2), n_(2), n_(3), l1(2), l2(1), _
    divide_string(v(1), v(2), True, False), temp_record, 0, 0, 0, 0, 0)
 'If combine_relation_with_relation_ > 1 Then
 ' Exit Function
 'End If
'End If
End If
End Function
Public Function combine_eline_with_three_line(ByVal el%, _
                ByVal start%, ByVal no_reduce As Byte) As Byte '10.10
Dim i%, k%, l%, no%
Dim n(1) As Integer
Dim m(2) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim th_l As line3_value_data0_type
Dim last_tn As Integer
Dim temp_record As total_record_type
Dim re As total_record_type
If Deline(el%).record_.no_reduce > 4 Then
 Exit Function
End If
Call add_conditions_to_record(eline_, el%, 0, 0, re.record_data.data0.condition_data)
re.record_data.data0.theorem_no = 1
For k% = 0 To 1
n(0) = k%
n(1) = (k% + 1) Mod 2
For l% = 0 To 2
m(0) = l%
m(1) = (l% + 1) Mod 3
m(2) = (l% + 2) Mod 3
th_l.poi(2 * m(0)) = Deline(el%).data(0).data0.poi(2 * n(0))
th_l.poi(2 * m(0) + 1) = Deline(el%).data(0).data0.poi(2 * n(0) + 1)
th_l.poi(2 * m(1)) = -1
Call search_for_line3_value(th_l, l%, n_(0), 1)
th_l.poi(2 * m(1)) = 30000
Call search_for_line3_value(th_l, l%, n_(1), 1)  '5.7
last_tn = 0
For i% = n_(0) + 1 To n_(1)
no% = line3_value(i%).data(0).record.data1.index.i(l%)
If line3_value(no%).record_.no_reduce < 4 And _
    no% > start% Then
   ' If is_two_record_related(eline_, el%, Deline(el%).data(0).record, _
        line3_value_, no%, line3_value(no%).data(0).record) = False Then
last_tn = last_tn + 1
ReDim Preserve tn(last_tn) As Integer
tn(last_tn) = no%
End If
'End If
Next i%
For i% = 1 To last_tn
no% = tn(i%)
temp_record = re
Call add_conditions_to_record(line3_value_, no%, 0, 0, temp_record.record_data.data0.condition_data)
combine_eline_with_three_line = set_three_line_value( _
  Deline(el%).data(0).data0.poi(2 * n(1)), Deline(el%).data(0).data0.poi(2 * n(1) + 1), _
   line3_value(no%).data(0).data0.poi(2 * m(1)), line3_value(no%).data(0).data0.poi(2 * m(1) + 1), _
    line3_value(no%).data(0).data0.poi(2 * m(2)), line3_value(no%).data(0).data0.poi(2 * m(2) + 1), _
  Deline(el%).data(0).data0.n(2 * n(1)), Deline(el%).data(0).data0.n(2 * n(1) + 1), _
   line3_value(no%).data(0).data0.n(2 * m(1)), line3_value(no%).data(0).data0.n(2 * m(1) + 1), _
    line3_value(no%).data(0).data0.n(2 * m(2)), line3_value(no%).data(0).data0.n(2 * m(2) + 1), _
  Deline(el%).data(0).data0.line_no(n(1)), line3_value(no%).data(0).data0.line_no(m(1)), _
   line3_value(no%).data(0).data0.line_no(m(2)), line3_value(no%).data(0).data0.para(m(0)), _
     line3_value(no%).data(0).data0.para(m(1)), line3_value(no%).data(0).data0.para(m(2)), _
      line3_value(no%).data(0).data0.value, temp_record, 0, no_reduce, 0)
If combine_eline_with_three_line > 0 Then
   Call set_level_(line3_value(no%).record_.no_reduce, 4)
 If combine_eline_with_three_line > 1 Then
  Exit Function
 End If
End If
Next i%
Next l%
Next k%
End Function
Public Function combine_relation_with_three_line_(ByVal ty1 As Byte, ByVal re%, _
     ty2 As Byte, ByVal t_l%, k%, l%) As Byte
Dim p1(5) As Integer
Dim n1(5) As Integer
Dim l1(2) As Integer
Dim p2(5) As Integer
Dim n2(5) As Integer
Dim l2(2) As Integer
Dim v(1) As String
Dim para(2) As String
Dim value$
Dim temp_record As total_record_type
Call add_conditions_to_record(ty1, re%, 0, 0, temp_record.record_data.data0.condition_data)
Call add_conditions_to_record(ty2, t_l%, 0, 0, temp_record.record_data.data0.condition_data)
     temp_record.record_data.data0.theorem_no = 1
Call read_point_and_ratio_from_relation(ty1, re%, k%, p1(), n1(), _
                                             l1(), v(0), v(1))
Call read_point_and_value_from_line_value(ty2, t_l%, l%, p2(), n2(), _
                                       l2(), para(), value$)
If v(0) <> "" Then
 combine_relation_with_three_line_ = set_three_line_value( _
  p1(2), p1(3), p2(2), p2(3), p2(4), p2(5), n1(2), n1(3), n2(2), n2(3), _
    n2(4), n2(5), l1(1), l2(1), l2(2), _
     time_string(v(0), para(0), True, False), para(1), para(2), _
       value$, temp_record, 0, 0, 0)
   Call set_record_no_reduce(ty2, t_l%, 0, 0, 255)
  If combine_relation_with_three_line_ > 1 Then
   Exit Function
  End If
End If
If v(1) <> "" Then
 combine_relation_with_three_line_ = set_three_line_value( _
  p1(4), p1(5), p2(2), p2(3), p2(4), p2(5), n1(4), n1(5), _
   n2(2), n2(3), n2(4), n2(5), l1(2), l2(1), l2(2), _
   time_string(v(1), para(0), True, False), para(1), para(2), _
       value$, temp_record, 0, 0, 0)
   Call set_record_no_reduce(ty2, t_l%, 0, 0, 255)
End If
End Function

Public Function combine_relation_with_three_line(ByVal l_r%, _
        ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, k%, l%, no%
Dim n(2) As Integer
Dim m(2) As Integer
Dim n_(1) As Integer
Dim s(1) As String
Dim tn() As Integer
Dim th_l As line3_value_data0_type
Dim last_tn As Integer
If Drelation(l_r%).record_.no_reduce > 4 Then
 Exit Function
End If
For k% = 0 To 2
Call read_ratio_from_Drelation(l_r%, k%, s(0), s(1))
n(0) = k%
n(1) = (k% + 1) Mod 3
n(2) = (k% + 2) Mod 3
For l% = 0 To 2
m(0) = l%
m(1) = (l% + 1) Mod 3
m(2) = (l% + 2) Mod 3
th_l.poi(2 * m(0)) = Drelation(l_r%).data(0).data0.poi(2 * n(0))
th_l.poi(2 * m(0) + 1) = Drelation(l_r%).data(0).data0.poi(2 * n(0) + 1)
th_l.poi(2 * m(1)) = -1
Call search_for_line3_value(th_l, l%, n_(0), 1)
th_l.poi(2 * m(1)) = 30000
Call search_for_line3_value(th_l, l%, n_(1), 1)  '5.7
last_tn = 0
For i% = n_(0) + 1 To n_(1)
no% = line3_value(i%).data(0).record.data1.index.i(l%)
If line3_value(no%).record_.no_reduce < 4 And _
    no% > start% Then
'If is_two_record_related(relation_, l_r%, Drelation(l_r%).data(0).record, _
      line3_value_, no%, line3_value(no%).data(0).record) = False Then
last_tn = last_tn + 1
ReDim Preserve tn(last_tn) As Integer
tn(last_tn) = no%
End If
'End If
Next i%
For i% = 1 To last_tn
no% = tn(i%)
combine_relation_with_three_line = combine_relation_with_three_line_( _
   relation_, l_r%, line3_value_, no%, n(0), m(0))
 If combine_relation_with_three_line > 1 Then
  Exit Function
 End If
Next i%
Next l%
Next k%
End Function

Public Function combine_mid_point_with_three_line(ByVal md%, _
             ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, no%
Dim n(5) As Integer
Dim m(2) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim th_l As line3_value_data0_type
Dim last_tn As Integer
If Dmid_point(md%).record_.no_reduce > 4 Then
 Exit Function
End If
For i% = 0 To 2
If i% = 0 Then
n(0) = 0
n(1) = 1
n(2) = 1
n(3) = 2
n(4) = 0
n(5) = 2
ElseIf i% = 1 Then
n(0) = 1
n(1) = 2
n(2) = 0
n(3) = 2
n(4) = 0
n(5) = 1
Else
n(0) = 0
n(1) = 2
n(2) = 0
n(3) = 1
n(4) = 1
n(5) = 2
End If
For j% = 0 To 2
m(0) = j%
m(1) = (j% + 1) Mod 3
m(2) = (j% + 2) Mod 3
th_l.poi(2 * m(0)) = Dmid_point(md%).data(0).data0.poi(n(0))
th_l.poi(2 * m(0) + 1) = Dmid_point(md%).data(0).data0.poi(n(1))
th_l.poi(2 * m(1)) = -1
Call search_for_line3_value(th_l, j%, n_(0), 1)
th_l.poi(2 * m(1)) = 30000
Call search_for_line3_value(th_l, j%, n_(1), 1)  '5.7
last_tn = 0
For k% = n_(0) + 1 To n_(1)
no% = line3_value(k%).data(0).record.data1.index.i(j%)
If line3_value(no%).record_.no_reduce < 4 And _
    no% > start% Then
'If is_two_record_related(midpoint_, md%, Dmid_point(md%).data(0).record, _
      line3_value_, no%, line3_value(no%).data(0).record) = False Then
last_tn = last_tn + 1
ReDim Preserve tn(last_tn) As Integer
tn(last_tn) = no%
End If
'End If
Next k%
For k% = 1 To last_tn
no% = tn(k%)
combine_mid_point_with_three_line = _
  combine_relation_with_three_line_(midpoint_, md%, line3_value_, _
   no%, i%, m(0))
 If combine_mid_point_with_three_line > 1 Then
  Exit Function
 End If
Next k%
Next j%
Next i%
End Function

Public Function combine_mid_point_with_eline(ByVal md%, _
          ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, k%, l%, no%
Dim n(5) As Integer
Dim m(2) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim e_l As eline_data0_type
Dim last_tn As Integer
Dim temp_record As total_record_type
If Dmid_point(md%).record_.no_reduce > 4 Then
 Exit Function
End If
For k% = 0 To 2
If k% = 0 Then
n(0) = 0
n(1) = 1
n(2) = 1
n(3) = 2
n(4) = 0
n(5) = 2
ElseIf k% = 1 Then
n(0) = 1
n(1) = 2
n(2) = 0
n(3) = 2
n(4) = 0
n(5) = 1
Else
n(0) = 0
n(1) = 2
n(2) = 0
n(3) = 1
n(4) = 1
n(5) = 2
End If
For l% = 0 To 1
m(0) = l%
m(1) = (l% + 1) Mod 2
e_l.poi(2 * m(0)) = Dmid_point(md%).data(0).data0.poi(n(0))
e_l.poi(2 * m(0) + 1) = Dmid_point(md%).data(0).data0.poi(n(1))
e_l.poi(2 * m(1)) = -1
Call search_for_eline(e_l, m(0), n_(0), 1)  '5.7
e_l.poi(2 * m(1)) = 30000
Call search_for_eline(e_l, m(0), n_(1), 1)
last_tn = 0
For i% = n_(0) + 1 To n_(1)
no% = Deline(i%).data(0).record.data1.index.i(m(0))  '5.7
If no% > start% And Deline(no%).record_.no_reduce < 4 Then
'If is_two_record_related(eline_, no%, Deline(no%).data(0).record, _
            midpoint_, md%, Dmid_point(md%).data(0).record) = False Then
last_tn = last_tn + 1
ReDim Preserve tn(last_tn) As Integer
tn(last_tn) = no%
End If
'End If
Next i%
For i% = 1 To last_tn
no% = tn(i%)
'***********************
If k% < 2 Then
 temp_record.record_data.data0.condition_data.condition_no = 2
 temp_record.record_data.data0.condition_data.condition(1).ty = eline_
 temp_record.record_data.data0.condition_data.condition(2).ty = midpoint_
 temp_record.record_data.data0.condition_data.condition(1).no = no%
 temp_record.record_data.data0.condition_data.condition(2).no = md%
 If l% = 0 Then
  If Deline(no%).data(0).data0.poi(2) = Dmid_point(md%).data(0).data0.poi(1) Then
   combine_mid_point_with_eline = set_dverti( _
     line_number0(Dmid_point(md%).data(0).data0.poi(0), Deline(no%).data(0).data0.poi(3), 0, 0), _
      line_number0(Dmid_point(md%).data(0).data0.poi(2), Deline(no%).data(0).data0.poi(3), 0, 0), _
        temp_record, 0, 0, False)
    If combine_mid_point_with_eline > 1 Then
       Exit Function
    End If
  ElseIf Deline(no%).data(0).data0.poi(3) = Dmid_point(md%).data(0).data0.poi(1) Then
   combine_mid_point_with_eline = set_dverti( _
     line_number0(Dmid_point(md%).data(0).data0.poi(0), Deline(no%).data(0).data0.poi(2), 0, 0), _
      line_number0(Dmid_point(md%).data(0).data0.poi(2), Deline(no%).data(0).data0.poi(2), 0, 0), _
        temp_record, 0, 0, False)
    If combine_mid_point_with_eline > 1 Then
       Exit Function
    End If
  End If
 ElseIf l% = 1 Then
  If Deline(no%).data(0).data0.poi(0) = Dmid_point(md%).data(0).data0.poi(1) Then
   combine_mid_point_with_eline = set_dverti( _
     line_number0(Dmid_point(md%).data(0).data0.poi(0), Deline(no%).data(0).data0.poi(1), 0, 0), _
      line_number0(Dmid_point(md%).data(0).data0.poi(2), Deline(no%).data(0).data0.poi(1), 0, 0), _
        temp_record, 0, 0, False)
    If combine_mid_point_with_eline > 1 Then
       Exit Function
    End If
  ElseIf Deline(no%).data(0).data0.poi(1) = Dmid_point(md%).data(0).data0.poi(1) Then
   combine_mid_point_with_eline = set_dverti( _
     line_number0(Dmid_point(md%).data(0).data0.poi(0), Deline(no%).data(0).data0.poi(0), 0, 0), _
      line_number0(Dmid_point(md%).data(0).data0.poi(2), Deline(no%).data(0).data0.poi(0), 0, 0), _
        temp_record, 0, 0, False)
    If combine_mid_point_with_eline > 1 Then
       Exit Function
    End If
  End If
 End If
End If
'***************************
combine_mid_point_with_eline = _
 combine_relation_with_relation_(midpoint_, md%, eline_, no%, _
    k%, l%)
If combine_mid_point_with_eline > 1 Then
 Exit Function
End If
Next i%
Next l%
Next k%
End Function

Public Function combine_general_string_with_general_string(ByVal g_s%, no_reduce) _
     As Byte
Dim i%, j%, k%, last_tn%, no%
Dim tn() As Integer
Dim n_(1) As Integer
Dim temp_record As total_record_type
Dim n(3) As Integer
Dim m(3) As Integer
Dim concl_no%
Dim tp(1) As Integer
Dim g_string As general_string_data_type
If general_string(g_s%).record_.no_reduce > 4 Then
 Exit Function
ElseIf general_string(g_s%).data(0).value <> "" And general_string(g_s%).record_.conclusion_no > 0 Then
 Exit Function
End If
'Call add_conditions_to_record(general_string_, g_s%, 0, 0, temp_record.record_data.data0.condition_data)
temp_record.record_data.data0.condition_data.condition_no = 2
 temp_record.record_data.data0.condition_data.condition(1).ty = general_string_
  temp_record.record_data.data0.condition_data.condition(2).ty = general_string_
   temp_record.record_data.data0.condition_data.condition(1).no = g_s%
    temp_record.record_data.data0.theorem_no = 1
For i% = 0 To 3
 n(0) = i%
  n(1) = (i% + 1) Mod 4
   n(2) = (i% + 2) Mod 4
    n(3) = (i% + 3) Mod 4
If general_string(g_s%).data(0).item(n(0)) > 0 Then
For j% = 0 To 3
 m(0) = j%
  m(1) = (j% + 1) Mod 4
   m(2) = (j% + 2) Mod 4
    m(3) = (j% + 3) Mod 4
g_string.item(m(0)) = general_string(g_s%).data(0).item(n(0))
 g_string.item(n(1)) = -1
Call search_for_general_string(g_string, j%, n_(0), 1)
 g_string.item(m(1)) = 30000
Call search_for_general_string(g_string, j%, n_(1), 1)
last_tn% = 0
'If general_string(g_s%).data(0).value = "" Then
'For k% = n_(0) + 1 To n_(1)
 'no% = general_string(k%).data(0).record.data1.index.i(j%)
 ' If no% > 0 And general_string(no%).data(0).value <> "" Then
'last_tn% = last_tn% + 1
'ReDim Preserve tn(last_tn%) As Integer
'tn(last_tn%) = no%
'End If
'Next k%
'Else
For k% = n_(0) + 1 To n_(1)
 no% = general_string(k%).data(0).record.data1.index.i(j%)
  If no% > 0 And g_s > no% Then
   If general_string(no%).record_.no_reduce <= 5 Then
   'If (g_s% > no%) Or general_string(g_s%).data(0).value = "" Or general_string(no%).data(0).value = ""
        '  Or general_string(no%).record_.conclusion_no = 0 Then
    'If (general_string(no%).data(0).value = "" And general_string(g_s%).data(0).value <> "") Or _
          (general_string(no%).data(0).value <> "" And general_string(g_s%).data(0).value = "") Or _
            (general_string(no%).record_.no_reduce < 4 And _
              general_string(g_s%).record_.no_reduce < 4) Then
'If is_two_record_related(general_string_, no%, general_string(no%).data(0).record, _
     general_string_, g_s%, general_string(g_s%).data(0).record) = False Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
End If
'End If
'End If
Next k%
'End If
For k% = 1 To last_tn%
no% = tn(k%)
If general_string(g_s%).data(0).value <> "" Or _
              general_string(no%).data(0).value <> "" Then
If general_string(no%).data(0).value = "" Then
temp_record.record_data.data0.condition_data.condition(1).no = g_s%
temp_record.record_data.data0.condition_data.condition(2).no = no%
Else 'if general_string(nog_s%).data(0)
temp_record.record_data.data0.condition_data.condition(2).no = g_s%
temp_record.record_data.data0.condition_data.condition(1).no = no%
End If
 If general_string(no%).data(0).value = "" Then
  temp_record.record_.conclusion_no = general_string(no%).record_.conclusion_no
  temp_record.record_.conclusion_ty = general_string(no%).record_.conclusion_ty
   concl_no% = general_string(no%).record_.conclusion_no
 ElseIf general_string(g_s%).data(0).value = "" Then
   concl_no% = general_string(g_s%).record_.conclusion_no
    temp_record.record_.conclusion_no = general_string(g_s%).record_.conclusion_no
    temp_record.record_.conclusion_ty = general_string(g_s%).record_.conclusion_ty
 ElseIf general_string(no%).data(0).value = "" And _
          general_string(g_s%).data(0).value = "" Then
           GoTo combine_general_string_with_general_string_next
 End If
combine_general_string_with_general_string = _
   combine_general_string_with_general_string_(general_string(g_s%).data(0).item(n(1)), _
     general_string(g_s%).data(0).item(n(2)), general_string(g_s%).data(0).item(n(3)), _
      general_string(g_s%).data(0).para(n(0)), general_string(g_s%).data(0).para(n(1)), _
       general_string(g_s%).data(0).para(n(2)), general_string(g_s%).data(0).para(n(3)), _
        general_string(no%).data(0).item(m(1)), general_string(no%).data(0).item(m(2)), _
         general_string(no%).data(0).item(m(3)), general_string(no%).data(0).para(m(0)), _
          general_string(no%).data(0).para(m(1)), general_string(no%).data(0).para(m(2)), _
           general_string(no%).data(0).para(m(3)), general_string(g_s%).data(0).value_, _
            general_string(no%).data(0).value_, concl_no%, temp_record, no_reduce)
Call set_record_no_reduce(general_string_, g_s%, _
     general_string_, no%, 255)
If combine_general_string_with_general_string > 1 Then
 Exit Function
End If
End If
combine_general_string_with_general_string_next:
Next k%
Next j%
End If
Next i%
End Function
Public Function combine_general_string_with_general_string_(ByVal it11%, _
    ByVal it12%, ByVal it13%, ByVal p10$, ByVal p11$, ByVal p12$, _
     ByVal p13$, ByVal it21%, ByVal it22%, ByVal it23%, ByVal p20$, _
      ByVal p21$, ByVal p22$, ByVal p23$, ByVal v1$, ByVal v2$, _
       concl_no%, re As total_record_type, no_reduce) As Byte
Dim i%, j%
Dim it(6) As Integer
Dim pA(6) As String
Dim rela_para As String
Dim v As String
Dim temp_record As total_record_type
Dim is_zero As Byte
temp_record = re
If v1$ <> "" And v2$ <> "" Then
 is_zero = 1
    Call set_level_(general_string(re.record_data.data0.condition_data.condition(2).no).record_.no_reduce, 4)
Else
 is_zero = 2
End If
it(0) = it11%
 it(1) = it12%
  it(2) = it13%
it(3) = it21%
 it(4) = it22%
  it(5) = it23%
it(6) = 0
If is_zero = 1 Then
pA(0) = determinant(p10$, p20$, p11$, "0")
 pA(1) = determinant(p10$, p20$, p12$, "0")
  pA(2) = determinant(p10$, p20$, p13$, "0")
pA(3) = determinant(p10$, p20$, "0", p21$)
 pA(4) = determinant(p10$, p20$, "0", p22$)
  pA(5) = determinant(p10$, p20$, "0", p23$)
v = determinant(p10$, p20$, v1$, v2$)
'End If
Else
 If v1$ <> "" Then
 rela_para = divide_string(p20$, p10$, True, False)
 If rela_para <> "F" Then
 pA(6) = time_string(v1$, rela_para, False, False)  'v2
 rela_para = time_string("-1", rela_para, True, False)
  pA(0) = time_string(p11$, rela_para, True, False)
   pA(1) = time_string(p12$, rela_para, True, False)
    pA(2) = time_string(p13$, rela_para, True, False)
 pA(3) = p21$
  pA(4) = p22$
   pA(5) = p23$
 Else
 Exit Function
 End If
v = ""
Else 'v2<>""
rela_para = divide_string(p10$, p20$, True, False)
If rela_para <> "F" Then
 pA(0) = p11$
  pA(1) = p12$
   pA(2) = p13$
 pA(6) = time_string(v2$, rela_para, False, False)  'v2
  rela_para = time_string("-1", rela_para, True, False)
 pA(3) = time_string(p21$, rela_para, True, False)
  pA(4) = time_string(p22$, rela_para, True, False)
   pA(5) = time_string(p23$, rela_para, True, False)
Else
 Exit Function
End If
v = ""
'concl_no% = general_string(re.record_data.data0.condition_data.condition(1).no).record_.conclusion_no
End If
End If
For i% = 0 To 2
 If pA(i%) = "0" Then
  it(i%) = 0
 End If
 For j% = 3 To 6
  If it(i%) = it(j%) Then
   pA(i%) = add_string(pA(i%), pA(j%), True, False)
    pA(j%) = "0"
     it(j%) = 0
   If pA(i%) = "0" Then
    it(i%) = 0
   End If
  End If
 Next j%
Next i%
Call remove_record_for_zero_para(pA(), it(), 5)
If pA(4) = "0" And it(4) = 0 Then
combine_general_string_with_general_string_ = _
 set_general_string(it(0), it(1), it(2), it(3), pA(0), pA(1), _
      pA(2), pA(3), v, concl_no, 0, is_zero, temp_record, 0, no_reduce)

End If
End Function
Public Function combine_item_with_dpoint_pair(it%, no_reduce As Byte) As Byte
Dim i%, j%, k%, no%, last_tn%
Dim n(1) As Integer
Dim m(3) As Integer
Dim dp As point_pair_data0_type
Dim tn() As Integer
Dim n_(1) As Integer
For i% = 0 To 1
n(0) = i%
 n(1) = (i% + 1) Mod 2
If item0(it%).data(0).poi(2 * n(0)) > 0 And item0(it%).data(0).poi(2 * n(0) + 1) > 0 Then
For j% = 0 To 3
 m(0) = j%
  m(1) = (j% + 1) Mod 4
dp.poi(2 * m(0)) = item0(it%).data(0).poi(2 * n(0))
dp.poi(2 * m(0) + 1) = item0(it%).data(0).poi(2 * n(0) + 1)
dp.poi(2 * m(1)) = -1
Call search_for_point_pair(dp, j%, n_(0), 1)
dp.poi(2 * m(1)) = 30000
Call search_for_point_pair(dp, j%, n_(1), 1)   '5.7
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
no% = Ddpoint_pair(k%).data(0).record.data1.index.i(j%)
If no% > 0 Then
 If Ddpoint_pair(no%).record_.no_reduce < 255 Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
 End If
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
combine_item_with_dpoint_pair = _
  combine_item_with_point_pair_(it%, no%, n(0), m(0), no_reduce)
If combine_item_with_dpoint_pair > 1 Then
 Exit Function
End If
Next k%
Next j%
End If
Next i%
End Function
Public Function solve_multi_varity_equations(ByVal A11$, ByVal A12$, ByVal A13$, _
                 ByVal A14$, ByVal b1$, ByVal A21$, ByVal A22$, ByVal A23$, _
                 ByVal A24$, ByVal b2$, s1$, S2$, s3$, v$) As Byte
Dim un_e$
  If InStr(1, A11$, "U", 0) > 0 Or InStr(1, A11$, "V", 0) > 0 Or _
     InStr(1, A12$, "U", 0) > 0 Or InStr(1, A11$, "V", 0) > 0 Or _
     InStr(1, A12$, "U", 0) > 0 Or InStr(1, A11$, "V", 0) > 0 Or _
     InStr(1, A14$, "U", 0) > 0 Or InStr(1, A11$, "V", 0) > 0 Or _
     InStr(1, A21$, "U", 0) > 0 Or InStr(1, A11$, "V", 0) > 0 Or _
     InStr(1, A22$, "U", 0) > 0 Or InStr(1, A11$, "V", 0) > 0 Or _
     InStr(1, A23$, "U", 0) > 0 Or InStr(1, A11$, "V", 0) > 0 Or _
     InStr(1, A24$, "U", 0) > 0 Or InStr(1, A11$, "V", 0) > 0 Then
     Exit Function
  End If
  s1$ = determinant(A11$, A21$, A12$, A22$)
  S2$ = determinant(A11$, A21$, A13$, A23$)
  s3$ = determinant(A11$, A21$, A14$, A24$)
  v$ = determinant(A11$, A21$, b1$, b2$)
  If InStr(1, s1$, "x", 0) > 0 Or _
     InStr(1, S2$, "x", 0) > 0 Or _
     InStr(1, s3$, "x", 0) > 0 Or _
     InStr(1, v$, "x", 0) > 0 Then
     solve_multi_varity_equations = 1
  Else 'If s1$ = "0" And S2$ = "0" And s3$ = "0" And V$ = "0" Then
     solve_multi_varity_equations = 2
  End If
End Function

Public Function combine_item_with_three_line_value_(ByVal it%, ByVal l3%, _
         k%, l%, no_reduce As Byte) As Byte
Dim temp_record0 As record_type0
Dim ite(1) As Integer
Dim tp(5) As Integer
Dim tn(5) As Integer
Dim tl(2) As Integer
Dim pA(2) As String
Dim n(2) As Integer
If line3_value(l3%).data(0).data0.value = "0" Then
n(0) = l%
n(1) = (l% + 1) Mod 3
n(2) = (l% + 2) Mod 3
pA(n(0)) = "1"
pA(n(1)) = time_string(line3_value(l3%).data(0).data0.para(n(0)), line3_value(l3%).data(0).data0.para(n(1)), _
              True, False)
pA(n(2)) = time_string(line3_value(l3%).data(0).data0.para(n(0)), line3_value(l3%).data(0).data0.para(n(2)), _
              True, False)
tp(0) = line3_value(l3%).data(0).data0.poi(2 * n(0))
tp(1) = line3_value(l3%).data(0).data0.poi(2 * n(0) + 1)
tp(2) = line3_value(l3%).data(0).data0.poi(2 * n(1))
tp(3) = line3_value(l3%).data(0).data0.poi(2 * n(1) + 1)
tp(4) = line3_value(l3%).data(0).data0.poi(2 * n(2))
tp(5) = line3_value(l3%).data(0).data0.poi(2 * n(2) + 1)
tn(0) = line3_value(l3%).data(0).data0.n(2 * n(0))
tn(1) = line3_value(l3%).data(0).data0.n(2 * n(0) + 1)
tn(2) = line3_value(l3%).data(0).data0.n(2 * n(1))
tn(3) = line3_value(l3%).data(0).data0.n(2 * n(1) + 1)
tn(4) = line3_value(l3%).data(0).data0.n(2 * n(2))
tn(5) = line3_value(l3%).data(0).data0.n(2 * n(2) + 1)
tl(0) = line3_value(l3%).data(0).data0.line_no(n(0))
tl(1) = line3_value(l3%).data(0).data0.line_no(n(1))
tl(2) = line3_value(l3%).data(0).data0.line_no(n(2))
   If pA(n(0)) = "1" Then
    If pA(n(1)) = "1" Or pA(n(1)) = "-1" Or pA(n(1)) = "@1" Then
     If pA(n(2)) = "1" Or pA(n(2)) = "-1" Or pA(n(2)) = "@1" Then
       If item0(it%).data(0).sig = "~" Then
        If k% = 0 Then
         combine_item_with_three_line_value_ = set_item0(tp(2), tp(3), 0, 0, "~", _
              tn(2), tn(3), 0, 0, tl(1), 0, "1", "1", "1", "", _
                "1", 0, record_data0.data0.condition_data, 0, ite(0), no_reduce, _
                   0, condition_data0, False)
          If combine_item_with_three_line_value_ > 1 Then
            Exit Function
         End If
        combine_item_with_three_line_value_ = set_item0(tp(4), tp(5), 0, 0, "~", _
              tn(4), tn(5), 0, 0, tl(2), 0, "1", "1", "1", "", _
                "1", 0, record_data0.data0.condition_data, 0, ite(1), no_reduce, _
                   0, condition_data0, False)
         If combine_item_with_three_line_value_ > 1 Then
            Exit Function
         End If
        Else
          Exit Function
        End If
       ElseIf item0(it%).data(0).sig = "*" Then
        If k% = 0 Then
         combine_item_with_three_line_value_ = set_item0(tp(2), tp(3), item0(it%).data(0).poi(2), _
             item0(it%).data(0).poi(3), "*", tn(2), tn(3), item0(it%).data(0).n(2), _
                    item0(it%).data(0).n(3), tl(1), item0(it%).data(0).line_no(1), _
                     "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, _
                       ite(0), no_reduce, 0, condition_data0, False)
         If combine_item_with_three_line_value_ > 1 Then
            Exit Function
         End If
         combine_item_with_three_line_value_ = set_item0(tp(4), tp(5), item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), _
                  "*", tn(4), tn(5), item0(it%).data(0).n(2), _
                    item0(it%).data(0).n(3), tl(2), item0(it%).data(0).line_no(1), _
                    "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, ite(1), _
                      no_reduce, 0, condition_data0, False)
         If combine_item_with_three_line_value_ > 1 Then
            Exit Function
         End If
        ElseIf k% = 1 Then
         combine_item_with_three_line_value_ = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), tp(2), tp(3), _
                  "*", item0(it%).data(0).n(0), item0(it%).data(0).n(1), tn(2), tn(3), _
                   item0(it%).data(0).line_no(0), tl(1), "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, _
                     0, ite(0), no_reduce, 0, condition_data0, False)
         If combine_item_with_three_line_value_ > 1 Then
            Exit Function
         End If
         combine_item_with_three_line_value_ = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), tp(4), tp(5), _
                  "*", item0(it%).data(0).n(0), item0(it%).data(0).n(1), tn(4), tn(5), _
                   item0(it%).data(0).line_no(0), tl(2), "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, _
                    0, ite(1), no_reduce, 0, condition_data0, False)
         If combine_item_with_three_line_value_ > 1 Then
            Exit Function
         End If
        Else
         Exit Function
        End If
       ElseIf item0(it%).data(0).sig = "/" Then
        If k% = 0 Then
         combine_item_with_three_line_value_ = set_item0(tp(2), tp(3), item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), _
                  "/", tn(2), tn(3), item0(it%).data(0).n(2), item0(it%).data(0).n(3), _
                   tl(1), item0(it%).data(0).line_no(1), "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, _
                    0, ite(0), no_reduce, 0, condition_data0, False)
         If combine_item_with_three_line_value_ > 1 Then
            Exit Function
         End If
         combine_item_with_three_line_value_ = set_item0(tp(4), tp(5), item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), _
                  "/", tn(4), tn(5), item0(it%).data(0).n(2), item0(it%).data(0).n(3), _
                   tl(2), item0(it%).data(0).line_no(1), "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, _
                    0, ite(1), no_reduce, 0, condition_data0, False)
          If combine_item_with_three_line_value_ > 1 Then
            Exit Function
         End If
       Else
         Exit Function
        End If
       End If
     End If
    End If
   End If
 End If
End Function
Public Function combine_item_with_general_string0(ByVal it%, ByVal trans_to_no%, _
           ByVal no%, k%, no_reduce As Byte) As Byte
Dim tn%, it1%, it2%, i%, j%, l%
Dim pA(4) As String
Dim ite(4) As Integer
Dim m(3) As Integer
Dim tv As String
Dim t_item As item0_data_type
Dim t_p(5) As Integer
Dim t_n(5) As Integer
Dim t_l(2) As Integer
Dim T_V(1) As String
Dim is_zero As Byte
Dim t_re_data As record_data_type
Dim temp_record As total_record_type
Dim temp_record_data As record_type0
If general_string(no%).data(0).value <> "" And general_string(no%).record_.conclusion_no > 0 Then
 '求值已有结果
 Exit Function
End If
   Call add_conditions_to_record(general_string_, no%, 0, 0, temp_record.record_data.data0.condition_data)
   If trans_to_no% > 0 Then
     Call add_record_to_record(item0(it%).data(0).record_for_trans.record(trans_to_no%).condition_data, _
      temp_record.record_data.data0.condition_data)
   End If
   temp_record.record_.conclusion_no = general_string(no%).record_.conclusion_no
    temp_record.record_.conclusion_ty = general_string(no%).record_.conclusion_ty
     temp_record.record_data.data0.theorem_no = 1
m(0) = k%
 m(1) = (k% + 1) Mod 4
  m(2) = (k% + 2) Mod 4
   m(3) = (k% + 3) Mod 4
pA(0) = general_string(no%).data(0).para(m(0))
pA(1) = general_string(no%).data(0).para(m(1))
pA(2) = general_string(no%).data(0).para(m(2))
pA(3) = general_string(no%).data(0).para(m(3))
pA(4) = "0"
ite(0) = general_string(no%).data(0).item(m(0))
ite(1) = general_string(no%).data(0).item(m(1))
ite(2) = general_string(no%).data(0).item(m(2))
ite(3) = general_string(no%).data(0).item(m(3))
ite(4) = 0
tv$ = general_string(no%).data(0).value
'***************************************************************
If trans_to_no% = -2 Then
 If general_string(no%).record_.no_reduce < 5 Then
 If (ite(1) > 0 And item0(ite(1)).data(0).no_reduce = True) Or ite(1) = 0 Then
   If (ite(2) > 0 And item0(ite(2)).data(0).no_reduce = True) Or ite(2) = 0 Then
     If (ite(3) > 0 And item0(ite(3)).data(0).no_reduce = True) Or ite(3) = 0 Then
       general_string(no%).record_.no_reduce = 5
        Exit Function
     End If
   End If
 End If
 End If
ElseIf trans_to_no% > 0 Then
 t_re_data.data0.condition_data = item0(it%).data(0).record_for_trans.record(trans_to_no%).condition_data
 Call add_record_to_record(t_re_data.data0.condition_data, temp_record.record_data.data0.condition_data)
If item0(it%).data(0).record_for_trans.record(trans_to_no%).to_no(1) = 0 Then
   If item0(it%).data(0).record_for_trans.record(trans_to_no%).para(1) <> "" Then
    pA(4) = time_string(item0(it%).data(0).record_for_trans.record(trans_to_no%).para(1), _
          pA(0), True, False)
    ite(4) = 0
   End If
   ite(0) = item0(it%).data(0).record_for_trans.record(trans_to_no%).to_no(0)
    pA(0) = time_string(pA(i%), item0(it%).data(0).record_for_trans.record(trans_to_no%).para(0), _
               True, False)
Else
   pA(4) = time_string(item0(it%).data(0).record_for_trans.record(trans_to_no%).para(1), _
          pA(0), True, False)
   ite(4) = item0(it%).data(0).record_for_trans.record(trans_to_no%).to_no(1)
   pA(0) = time_string(item0(it%).data(0).record_for_trans.record(trans_to_no%).para(0), _
          pA(0), True, False)
   ite(0) = item0(it%).data(0).record_for_trans.record(trans_to_no%).to_no(0)
End If
ElseIf trans_to_no% = -1 And item0(it%).data(0).value <> "" Then ' if =it%
ite(0) = 0
 If InStr(1, pA(0), "V", 0) > 0 Or InStr(1, pA(0), "U", 0) > 0 Then
    Exit Function
 End If
 pA(0) = time_string(pA(0), item0(it%).data(0).value, True, False)
  temp_record.record_data.data0.condition_data.condition_no = 1
   temp_record.record_data.data0.condition_data.condition(1).ty = general_string_
    temp_record.record_data.data0.condition_data.condition(1).no = no%
     temp_record.record_data.data0.theorem_no = 1
      t_re_data.data0.condition_data = item0(it%).data(0).record_for_value.data0.condition_data
       Call add_conditions_to_record(item0_, it%, 0, 0, _
                         temp_record.record_data.data0.condition_data)
Else
 Exit Function
End If
   For i% = 0 To 3
    For l% = i% + 1 To 4
     If ite(i%) = ite(l%) Then
       pA(i%) = add_string(pA(i%), pA(l%), True, False)
      For j% = l% To 3
       pA(j%) = pA(j% + 1)
       ite(j%) = ite(j% + 1)
      Next j%
       If pA(i%) = "0" Then
        ite(i%) = 0
       End If
       pA(4) = "0"
        ite(4) = 0
     End If
    Next l%
   Next i%
     Call remove_record_for_zero_para(pA(), ite(), 4)
  If tv$ <> "" Then
    For i% = 0 To 4
     If ite(i%) = 0 Then
      tv$ = minus_string(tv$, pA(i%), True, False)
      pA(i%) = "0"
       For l% = i% To 3
        ite(l%) = ite(l% + 1)
        pA(l%) = pA(l% + 1)
       Next l%
        ite(4) = 0
        pA(4) = "0"
     End If
    Next i%
  End If
  '********************
  If ite(4) = 0 And pA(4) = "0" Then
   If ite(0) = 0 And ite(1) = 0 And ite(2) = 0 And ite(3) = 0 Then
    If general_string(no%).data(0).value = "" Then
     Call add_record_to_record(t_re_data.data0.condition_data, general_string(no%).data(0).record.data0.condition_data)
     Call set_level(general_string(no%).data(0).record.data0.condition_data)
      If ite(1) = 0 And ite(2) = 0 And ite(3) = 0 Then
        general_string(no%).data(0).value = add_string(pA(0), pA(1), False, False)
        general_string(no%).data(0).value = add_string(general_string(no%).data(0).value, pA(2), False, False)
        general_string(no%).data(0).value = add_string(general_string(no%).data(0).value, pA(3), False, False)
        pA(0) = "0"
        pA(1) = "0"
        pA(2) = "0"
        pA(3) = "0"
      End If
      combine_item_with_general_string0 = is_con_general_string(no%)
     If combine_item_with_general_string0 > 1 Then
      Exit Function
     End If
    Else
     general_string(no%).record_.no_reduce = 255
      Exit Function
    End If
   Else
     If tv$ <> "" Then
       If tv$ = "0" And ite(2) = 0 And ite(0) > 0 And ite(1) > 0 Then
          If item0(ite(0)).data(0).sig = "*" And item0(ite(1)).data(0).sig = "*" Then
           If time_string("-1", pA(0), True, False) <> pA(1) Then
            Exit Function
           Else
            combine_item_with_general_string0 = _
             set_dpoint_pair(item0(ite(0)).data(0).poi(0), item0(ite(0)).data(0).poi(1), _
               item0(ite(1)).data(0).poi(0), item0(ite(1)).data(0).poi(1), item0(ite(1)).data(0).poi(2), _
                item0(ite(1)).data(0).poi(3), item0(ite(0)).data(0).poi(2), item0(ite(0)).data(0).poi(3), _
                 item0(ite(0)).data(0).n(0), item0(ite(0)).data(0).n(1), item0(ite(1)).data(0).n(0), _
                  item0(ite(1)).data(0).n(1), item0(ite(1)).data(0).n(2), item0(ite(1)).data(0).n(3), _
                   item0(ite(0)).data(0).n(2), item0(ite(0)).data(0).n(3), item0(ite(0)).data(0).line_no(0), _
                    item0(ite(1)).data(0).line_no(0), item0(ite(1)).data(0).line_no(1), item0(ite(0)).data(0).line_no(1), _
                     0, temp_record, True, 0, 0, 0, no_reduce, False)
             Exit Function
           End If
        End If
      End If
     End If
     End If
     'Call add_record_to_record(item0(it%).data(0).record_for_trans.record(trans_to_no%).condition_data, temp_record.record_data.data0.condition_data)
     combine_item_with_general_string0 = set_general_string(ite(0), _
       ite(1), ite(2), ite(3), pA(0), pA(1), pA(2), pA(3), tv$, _
        general_string(no%).record_.conclusion_no, 0, _
          is_zero, temp_record, 0, no_reduce)
      If combine_item_with_general_string0 > 0 Then
       If general_string(no%).data(0).value <> "" Then
        Call set_level_(general_string(no%).record_.no_reduce, 4)
       End If
       If combine_item_with_general_string0 > 1 Then
        Exit Function
       End If
      End If
'     End If
  ' End If
  End If
'***
End Function

Public Function combine_item_with_midpoint(I0 As Integer, no_reduce As Byte) As Byte
Dim i%, j%, no%, k%, tn1%
Dim v1$, v2$
Dim n(2) As Integer
Dim m(5) As Integer
Dim tn0() As Integer
Dim it As item0_data_type
Dim tv As String
Dim t_it As Integer
Dim md  As mid_point_data0_type
For i% = 0 To 2
 n(0) = i%
  n(1) = (i% + 1) Mod 3
   n(2) = (i% + 2) Mod 3
 If item0(I0).data(0).poi(2 * n(0)) > 0 And item0(I0).data(0).poi(2 * n(0) + 1) > 0 Then
 For j% = 0 To 2
  If j% = 0 Then
    m(0) = 0
     m(1) = 1
      m(2) = 1
    m(3) = 2
     m(4) = 0
      m(5) = 2
   ElseIf j% = 1 Then
    m(0) = 1
     m(1) = 2
      m(2) = 0
    m(3) = 2
     m(4) = 0
      m(5) = 1
   ElseIf j% = 2 Then
    m(0) = 0
     m(1) = 2
      m(2) = 0
    m(3) = 1
     m(4) = 1
      m(5) = 2
   End If
 md.poi(m(0)) = item0(I0).data(0).poi(2 * n(0))
 md.poi(m(1)) = item0(I0).data(0).poi(2 * n(0) + 1)
 If search_for_mid_point(md, j%, no%, 2) Then  '5.7原j%+3
    If Dmid_point(no%).record_.no_reduce < 255 Then
    combine_item_with_midpoint = _
      combine_relation_with_item_(midpoint_, no%, I0, _
       j%, i%, no_reduce)
  If combine_item_with_midpoint > 1 Then
   Exit Function
  End If
 End If
 End If
 Next j%
 End If
Next i%
End Function

Public Function combine_mid_point_with_item(md%, no_reduce As Byte) As Integer
Dim i%, j%, k%, no%, tn1%
Dim n_(1) As Integer
Dim n0(2) As Integer
Dim n(5) As Integer
Dim m(2) As Integer
Dim it As item0_data_type
Dim tn() As Integer
Dim last_tn%
For i% = 0 To 2
 If i% = 0 Then
 n0(0) = 0
 n0(1) = 1
 n0(2) = 2
 n(0) = 0
  n(1) = 1
   n(2) = 1
 n(3) = 2
  n(4) = 0
   n(5) = 2
 ElseIf i% = 1 Then
 n0(0) = 1
 n0(1) = 2
 n0(2) = 0
 n(0) = 1
  n(1) = 2
   n(2) = 0
 n(3) = 2
  n(4) = 0
   n(5) = 1
 ElseIf i% = 2 Then
  n0(0) = 2
  n0(1) = 0
  n0(2) = 1
  n(0) = 0
  n(1) = 2
   n(2) = 0
 n(3) = 1
  n(4) = 1
   n(5) = 2
 End If
For j% = 0 To 2
 m(0) = j%
   m(1) = (j% + 1) Mod 3
    m(2) = (j% + 2) Mod 3
it.poi(2 * m(0)) = Dmid_point(md%).data(0).data0.poi(n(0))
it.poi(2 * m(0) + 1) = Dmid_point(md%).data(0).data0.poi(n(1))
it.poi(2 * m(1)) = -1
Call search_for_item0(it, m(0), n_(0), 1)
it.poi(2 * m(1)) = 30000
Call search_for_item0(it, m(0), n_(1), 1)   '5.7
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
 no% = item0(k%).data(0).index(m(0))
 If no% > 0 Then
    last_tn% = last_tn% + 1
     ReDim Preserve tn(last_tn%) As Integer
    tn(last_tn%) = no%
  End If
Next k%
 For k% = 1 To last_tn%
 no% = tn(k%)
     combine_mid_point_with_item = _
      combine_relation_with_item_(midpoint_, md%, no%, i%, j%, no_reduce)
     If combine_mid_point_with_item > 1 Then
      Exit Function
     End If
Next k%
Next j%
Next i%
End Function
Public Function combine_item_with_point_pair_(ByVal it%, _
         ByVal p%, k%, l%, no_reduce As Byte) As Byte
Dim tn%
Dim n(3) As Integer
Dim m(1) As Integer
Dim temp_record0 As condition_data_type
Dim ite As item0_data_type
If k% < 2 Then
temp_record0.condition_no = 1
temp_record0.condition(1).ty = dpoint_pair_
temp_record0.condition(1).no = p%
If l% = 0 Then
n(0) = 0
n(1) = 1
n(2) = 2
n(3) = 3
ElseIf l% = 1 Then
n(0) = 1
n(1) = 0
n(2) = 3
n(3) = 2
ElseIf l% = 2 Then
n(0) = 2
n(1) = 3
n(2) = 0
n(3) = 1
Else
n(0) = 3
n(1) = 2
n(2) = 1
n(3) = 0
End If
m(0) = k%
m(1) = (k% + 1) Mod 2
ite = item0(it%).data(0)
If ite.sig = "*" Then
    If ite.poi(2 * m(1)) = Ddpoint_pair(p%).data(0).data0.poi(2 * n(3)) And _
         ite.poi(2 * m(1) + 1) = Ddpoint_pair(p%).data(0).data0.poi(2 * n(3) + 1) Then
      tn% = 0
       combine_item_with_point_pair_ = set_item0( _
        Ddpoint_pair(p%).data(0).data0.poi(2 * n(1)), _
         Ddpoint_pair(p%).data(0).data0.poi(2 * n(1) + 1), _
          Ddpoint_pair(p%).data(0).data0.poi(2 * n(2)), _
           Ddpoint_pair(p%).data(0).data0.poi(2 * n(2) + 1), _
        "*", Ddpoint_pair(p%).data(0).data0.n(2 * n(1)), _
         Ddpoint_pair(p%).data(0).data0.n(2 * n(1) + 1), _
          Ddpoint_pair(p%).data(0).data0.n(2 * n(2)), _
           Ddpoint_pair(p%).data(0).data0.n(2 * n(2) + 1), _
            Ddpoint_pair(p%).data(0).data0.line_no(n(1)), _
             Ddpoint_pair(p%).data(0).data0.line_no(n(2)), "1", "1", "1", _
              "", "1", 0, temp_record0, it%, tn%, no_reduce, 0, condition_data0, False)
       If combine_item_with_point_pair_ > 1 Then
        Exit Function
       End If
       combine_item_with_point_pair_ = add_new_item_to_item(tn%, 0, _
          "1", "0", it%, temp_record0)
       If combine_item_with_point_pair_ > 1 Then
        Exit Function
       End If
    End If
ElseIf ite.sig = "/" Then
 If k% = 0 Then
    If ite.poi(2) = Ddpoint_pair(p%).data(0).data0.poi(2 * n(1)) And _
         ite.poi(3) = Ddpoint_pair(p%).data(0).data0.poi(2 * n(1) + 1) Then
       combine_item_with_point_pair_ = set_item0( _
        Ddpoint_pair(p%).data(0).data0.poi(2 * n(2)), _
         Ddpoint_pair(p%).data(0).data0.poi(2 * n(2) + 1), _
          Ddpoint_pair(p%).data(0).data0.poi(2 * n(3)), _
           Ddpoint_pair(p%).data(0).data0.poi(2 * n(3) + 1), _
        "/", Ddpoint_pair(p%).data(0).data0.n(2 * n(2)), _
         Ddpoint_pair(p%).data(0).data0.n(2 * n(2) + 1), _
          Ddpoint_pair(p%).data(0).data0.n(2 * n(3)), _
           Ddpoint_pair(p%).data(0).data0.n(2 * n(3) + 1), _
            Ddpoint_pair(p%).data(0).data0.line_no(n(2)), _
             Ddpoint_pair(p%).data(0).data0.line_no(n(3)), "1", "1", "1", _
               "", "1", 0, temp_record0, it%, tn%, no_reduce, 0, condition_data0, False)
       If combine_item_with_point_pair_ > 1 Then
        Exit Function
       End If
        combine_item_with_point_pair_ = add_new_item_to_item(tn%, 0, _
          "1", "0", it%, temp_record0)
       If combine_item_with_point_pair_ > 1 Then
        Exit Function
       End If
     ElseIf ite.poi(2) = Ddpoint_pair(p%).data(0).data0.poi(2 * n(2)) And _
         ite.poi(3) = Ddpoint_pair(p%).data(0).data0.poi(2 * n(2) + 1) Then
      combine_item_with_point_pair_ = _
       set_item0(Ddpoint_pair(p%).data(0).data0.poi(2 * n(1)), _
         Ddpoint_pair(p%).data(0).data0.poi(2 * n(1) + 1), _
          Ddpoint_pair(p%).data(0).data0.poi(2 * n(3)), _
           Ddpoint_pair(p%).data(0).data0.poi(2 * n(3) + 1), _
         "/", Ddpoint_pair(p%).data(0).data0.n(2 * n(1)), _
           Ddpoint_pair(p%).data(0).data0.n(2 * n(1) + 1), _
            Ddpoint_pair(p%).data(0).data0.n(2 * n(3)), _
             Ddpoint_pair(p%).data(0).data0.n(2 * n(3) + 1), _
              Ddpoint_pair(p%).data(0).data0.line_no(n(1)), _
               Ddpoint_pair(p%).data(0).data0.line_no(n(3)), "1", "1", _
                 "1", "", "1", 0, temp_record0, it%, tn%, no_reduce, 0, condition_data0, False)
       If combine_item_with_point_pair_ > 1 Then
         Exit Function
       End If
       '  combine_item_with_point_pair_ = add_new_item_to_item(tn%, 0, _
          "1", "0", it%, temp_record0)
       '  If combine_item_with_point_pair_ > 1 Then
       '  Exit Function
       '  End If
       End If
  ElseIf k% = 1 Then
    If ite.poi(0) = Ddpoint_pair(p%).data(0).data0.poi(2 * n(1)) And _
         ite.poi(1) = Ddpoint_pair(p%).data(0).data0.poi(2 * n(1) + 1) Then
      combine_item_with_point_pair_ = set_item0( _
        Ddpoint_pair(p%).data(0).data0.poi(2 * n(3)), _
         Ddpoint_pair(p%).data(0).data0.poi(2 * n(3) + 1), _
          Ddpoint_pair(p%).data(0).data0.poi(2 * n(2)), _
           Ddpoint_pair(p%).data(0).data0.poi(2 * n(2) + 1), _
            "/", Ddpoint_pair(p%).data(0).data0.n(2 * n(3)), _
             Ddpoint_pair(p%).data(0).data0.n(2 * n(3) + 1), _
              Ddpoint_pair(p%).data(0).data0.n(2 * n(2)), _
               Ddpoint_pair(p%).data(0).data0.n(2 * n(2) + 1), _
                Ddpoint_pair(p%).data(0).data0.line_no(n(3)), _
                 Ddpoint_pair(p%).data(0).data0.line_no(n(2)), "1", "1", _
                  "1", "", "1", 0, temp_record0, it%, _
                     tn%, no_reduce, 0, condition_data0, False)
       If combine_item_with_point_pair_ > 1 Then
        Exit Function
       End If
 '       combine_item_with_point_pair_ = add_new_item_to_item(tn%, 0, _
          "1", "0", it%, temp_record0)
 '      If combine_item_with_point_pair_ > 1 Then
 '        Exit Function
 '      End If
      ElseIf ite.poi(0) = Ddpoint_pair(p%).data(0).data0.poi(2 * n(2)) And _
         ite.poi(1) = Ddpoint_pair(p%).data(0).data0.poi(2 * n(2) + 1) Then
      combine_item_with_point_pair_ = set_item0( _
        Ddpoint_pair(p%).data(0).data0.poi(2 * n(3)), _
         Ddpoint_pair(p%).data(0).data0.poi(2 * n(3) + 1), _
          Ddpoint_pair(p%).data(0).data0.poi(2 * n(1)), _
           Ddpoint_pair(p%).data(0).data0.poi(2 * n(1) + 1), _
            "/", Ddpoint_pair(p%).data(0).data0.n(2 * n(3)), _
             Ddpoint_pair(p%).data(0).data0.n(2 * n(3) + 1), _
              Ddpoint_pair(p%).data(0).data0.n(2 * n(1)), _
               Ddpoint_pair(p%).data(0).data0.n(2 * n(1) + 1), _
                Ddpoint_pair(p%).data(0).data0.line_no(n(3)), _
                 Ddpoint_pair(p%).data(0).data0.line_no(n(1)), "1", "1", _
                  "1", "", "1", 0, temp_record0, it%, _
                    tn%, no_reduce, 0, condition_data0, False)
       If combine_item_with_point_pair_ > 1 Then
        Exit Function
       End If
   '    combine_item_with_point_pair_ = add_new_item_to_item(tn%, 0, _
          "1", "0", it%, temp_record0)
       End If
       End If
     End If
 End If
End Function
Public Function combine_item_with_relation(ByVal I0 As Integer, no_reduce As Byte) As Byte
Dim i%, j%, k%, no%
Dim n(2) As Integer
Dim m(2) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim last_tn%
Dim re As relation_data0_type
For i% = 0 To 2
 n(0) = i%
  n(1) = (i% + 1) Mod 3
   n(2) = (i% + 2) Mod 3
If item0(I0).data(0).poi(2 * n(0)) > 0 And item0(I0).data(0).poi(2 * n(0) + 1) Then
 For j% = 0 To 2
  m(0) = j%
   m(1) = (j% + 1) Mod 3
    m(2) = (j% + 2) Mod 3
 re.poi(2 * m(0)) = item0(I0).data(0).poi(2 * n(0))
 re.poi(2 * m(0) + 1) = item0(I0).data(0).poi(2 * n(0) + 1)
 re.poi(2 * m(1)) = -1
 Call search_for_relation(re, m(0), n_(0), 1)  '5.7
 re.poi(2 * m(1)) = 30000
 Call search_for_relation(re, m(0), n_(1), 1)
 last_tn% = 0
 For k% = n_(0) + 1 To n_(1)
 no% = Drelation(k%).data(0).record.data1.index.i(m(0))
 If no% > 0 Then
 If Drelation(no%).record_.no_reduce < 255 Then
 last_tn% = last_tn% + 1
 ReDim Preserve tn(last_tn%) As Integer
 tn(last_tn%) = no%
 End If
 End If
 Next k%
  For k% = 1 To last_tn%
   no% = tn(k%)
   combine_item_with_relation = combine_relation_with_item_(relation_, no%, _
    I0, j%, i%, no_reduce)
  If combine_item_with_relation > 1 Then
   Exit Function
  End If
  Next k%
  Next j%
 End If
Next i%
End Function




Public Function combine_two_triangle(ByVal triA1%, ByVal triA2%, _
   otriA1%, otriA2%, otriA3%, poly4_no%) As Byte '负
Dim i%, j%, k%, l% ' 0,和,1,1-2,-1 2-1,10并成一个四边形
Dim p1(2) As Integer
Dim p2(2) As Integer
Dim p3(4) As Integer
Dim tn1(3) As Integer
Dim tn2(3) As Integer
Dim tn(1) As Integer
Dim l1(1) As Integer
Dim l2(1) As Integer
Dim ty As Byte
Dim triangle_data(1) As triangle_data0_type
Dim is_no_initial As Integer
Dim c_data As condition_data_type
triangle_data(0) = triangle(triA1%).data(0)
triangle_data(1) = triangle(triA2%).data(0)
For i% = 0 To 2
 For j% = 0 To 2 '两个顶点同
  If triangle_data(0).poi(i%) = triangle_data(1).poi(j%) Then
   If triangle_data(0).poi((i% + 1) Mod 3) = _
        triangle_data(1).poi((j% + 1) Mod 3) Then
   p1(0) = triangle_data(0).poi((i% + 1) Mod 3)
    p2(0) = triangle_data(1).poi((j% + 1) Mod 3)
   p1(1) = triangle_data(0).poi((i% + 2) Mod 3)
    p2(1) = triangle_data(1).poi((j% + 2) Mod 3)
   p1(2) = triangle_data(0).poi(i%)
    p2(2) = triangle_data(1).poi(j%)
   ElseIf triangle_data(0).poi((i% + 2) Mod 3) = _
        triangle_data(1).poi((j% + 2) Mod 3) Then
   p1(0) = triangle_data(0).poi((i% + 2) Mod 3)
    p2(0) = triangle_data(1).poi((j% + 2) Mod 3)
   p1(1) = triangle_data(0).poi((i% + 1) Mod 3)
    p2(1) = triangle_data(1).poi((j% + 1) Mod 3)
   p1(2) = triangle_data(0).poi(i%)
    p2(2) = triangle_data(1).poi(j%)
   ElseIf triangle_data(0).poi((i% + 1) Mod 3) = _
        triangle_data(1).poi((j% + 2) Mod 3) Then
   p1(0) = triangle_data(0).poi((i% + 1) Mod 3)
    p2(0) = triangle_data(1).poi((j% + 2) Mod 3)
   p1(1) = triangle_data(0).poi((i% + 2) Mod 3)
    p2(1) = triangle_data(1).poi((j% + 1) Mod 3)
   p1(2) = triangle_data(0).poi(i%)
    p2(2) = triangle_data(1).poi(j%)
   ElseIf triangle_data(0).poi((i% + 2) Mod 3) = _
        triangle_data(1).poi((j% + 1) Mod 3) Then
   p1(0) = triangle_data(0).poi((i% + 2) Mod 3)
    p2(0) = triangle_data(1).poi((j% + 1) Mod 3)
   p1(1) = triangle_data(0).poi((i% + 1) Mod 3)
    p2(1) = triangle_data(1).poi((j% + 2) Mod 3)
   p1(2) = triangle_data(0).poi(i%)
    p2(2) = triangle_data(1).poi(j%)
   Else
    combine_two_triangle = 0
     otriA1% = triA1%
      otriA2% = triA2%
     Exit Function
   End If
   l1(0) = line_number0(p1(0), p1(1), tn1(0), tn1(1))
   l1(1) = line_number0(p1(2), p1(1), tn1(2), tn1(3))
   l2(0) = line_number0(p2(0), p2(1), tn2(0), tn2(1))
   l2(1) = line_number0(p2(2), p2(1), tn2(2), tn2(3))
   If l1(0) = l2(0) Then '两线
    Call arrange_four_point(p1(0), p1(1), p2(0), p2(1), _
      tn1(0), tn1(1), tn2(0), tn2(1), l1(0), l2(0), _
       p3(0), p3(1), p3(2), p3(3), 0, 0, 0, 0, 0, _
        0, 0, 0, 0, 0, 0, combine_two_triangle, c_data, is_no_initial)
         p3(4) = p1(2)
   ElseIf l1(1) = l2(1) Then
     Call arrange_four_point(p1(1), p1(2), p2(1), p2(2), _
       tn1(3), tn1(2), tn2(3), tn2(2), l1(1), l2(1), _
        p3(0), p3(1), p3(2), p3(3), 0, 0, 0, 0, 0, _
        0, 0, 0, 0, 0, 0, combine_two_triangle, c_data, is_no_initial)
        p3(4) = p1(0)
   Else
    If (angle_number(p1(0), p1(1), p1(2), "", 0) > 0 And _
        angle_number(p1(0), p1(1), p2(2), "", 0) < 0) Or _
       (angle_number(p1(0), p1(1), p1(2), "", 0) < 0 And _
        angle_number(p1(0), p1(1), p2(2), "", 0) > 0) Then
         poly4_no% = polygon4_number(p1(0), p1(1), p1(2), p2(2), 0)
          combine_two_triangle = 0
           otriA3% = 0
            If triA1% > triA2% Then
             otriA1% = triA1%
              otriA2% = triA2%
               combine_two_triangle = 3
            Else
             otriA1% = triA2%
              otriA2% = triA1%
               combine_two_triangle = 5
            End If
    Else
     combine_two_triangle = 0
      otriA1% = triA1%
       otriA2% = triA2%
    End If
     Exit Function
   End If
    If combine_two_triangle = 3 Or combine_two_triangle = 5 Then
    otriA3% = triangle_number(p3(0), p3(3), p3(4), 0, 0, 0, 0, _
            0, 0, 0)
     If combine_two_triangle = 3 Then
      otriA1% = triA1%
       otriA2% = triA2%
     Else
      otriA1% = triA2%
       otriA2% = triA1%
     End If
    ElseIf combine_two_triangle = 4 Or combine_two_triangle = 6 Then
    otriA1% = triangle_number(p3(0), p3(2), p3(4), 0, 0, 0, 0, _
                  0, 0, 0)
     If combine_two_triangle = 4 Then
      otriA2% = triA2%
       otriA3% = triA1%
     Else
      otriA2% = triA1%
       otriA3% = triA2%
     End If
    ElseIf combine_two_triangle = 7 Or combine_two_triangle = 8 Then
    otriA2% = triangle_number(p3(2), p3(3), p3(4), 0, 0, 0, 0, _
               0, 0, 0)
     If combine_two_triangle = 7 Then
      otriA1% = triA1%
       otriA3% = triA2%
     Else
      otriA1% = triA2%
       otriA3% = triA1%
     End If
    Else
      otriA1% = triA1%
       otriA2% = triA2%
    End If
     Exit Function
    End If
   Next j%
Next i%
If combine_two_triangle = 0 Then
     otriA1% = triA1%
      otriA2% = triA2%
End If
End Function

Public Function combine_line_value_with_dpoint_pair_(ByVal l%, _
   ByVal dp%, n1%, n2%, n3%, n4%, no_reduce As Byte) As Byte
Dim temp_record As total_record_type
Dim re As record_data_type
Dim tn%, it%
Dim t_dp As point_pair_data0_type
re.data0.condition_data.condition_no = 2
re.data0.condition_data.condition(1).ty = dpoint_pair_
re.data0.condition_data.condition(2).ty = line_value_
re.data0.condition_data.condition(1).no = dp%
re.data0.condition_data.condition(2).no = l%
re.data0.theorem_no = 1
     If is_line_value(Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%), _
         Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1), _
          Ddpoint_pair(dp%).data(0).data0.n(2 * n2%), _
             Ddpoint_pair(dp%).data(0).data0.n(2 * n2% + 1), _
                Ddpoint_pair(dp%).data(0).data0.line_no(n2%), _
            "", tn%, -1000, 0, 0, 0, line_value_data0) = 1 Then
        temp_record.record_data = re
        Call add_conditions_to_record(line_value_, tn%, 0, 0, temp_record.record_data.data0.condition_data)
       combine_line_value_with_dpoint_pair_ = _
        combine_relation_with_dpoint_pair00_(dp%, _
          n1%, n2%, n3%, n4%, divide_string(line_value(l%).data(0).data0.value, _
           line_value(tn%).data(0).data0.value, True, False), temp_record.record_data)
        If combine_line_value_with_dpoint_pair_ > 1 Then
         Exit Function
        End If
     ElseIf is_line_value(Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%), _
      Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1), _
         Ddpoint_pair(dp%).data(0).data0.n(2 * n3%), _
          Ddpoint_pair(dp%).data(0).data0.n(2 * n3% + 1), _
           Ddpoint_pair(dp%).data(0).data0.line_no(n3%), _
            "", tn%, -1000, 0, 0, 0, _
             line_value_data0) = 1 Then
        temp_record.record_data = re
        Call add_conditions_to_record(line_value_, tn%, 0, 0, temp_record.record_data.data0.condition_data)
       combine_line_value_with_dpoint_pair_ = _
        combine_relation_with_dpoint_pair00_(dp%, _
          n1%, n3%, n2%, n4%, divide_string(line_value(l%).data(0).data0.value, _
           line_value(tn%).data(0).data0.value, True, False), temp_record.record_data)
        If combine_line_value_with_dpoint_pair_ > 1 Then
         Exit Function
        End If
     ElseIf is_line_value(Ddpoint_pair(dp%).data(0).data0.poi(2 * n4%), _
      Ddpoint_pair(dp%).data(0).data0.poi(2 * n4% + 1), _
         Ddpoint_pair(dp%).data(0).data0.n(2 * n4%), _
          Ddpoint_pair(dp%).data(0).data0.n(2 * n4% + 1), _
           Ddpoint_pair(dp%).data(0).data0.line_no(n4%), _
             "", tn%, -1000, 0, 0, 0, _
              line_value_data0) = 1 Then
       temp_record.record_data = re
        Call add_conditions_to_record(line_value_, tn%, 0, 0, temp_record.record_data.data0.condition_data)
     combine_line_value_with_dpoint_pair_ = set_item0(Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%), _
       Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1), _
        Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%), _
         Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1), _
          "*", Ddpoint_pair(dp%).data(0).data0.n(2 * n2%), _
            Ddpoint_pair(dp%).data(0).data0.n(2 * n2% + 1), _
             Ddpoint_pair(dp%).data(0).data0.n(2 * n3%), _
              Ddpoint_pair(dp%).data(0).data0.n(2 * n3% + 1), _
               Ddpoint_pair(dp%).data(0).data0.line_no(n2%), _
                Ddpoint_pair(dp%).data(0).data0.line_no(n3%), "1", "1", _
                 "1", "", "1", 0, temp_record.record_data.data0.condition_data, _
                   0, it%, no_reduce, 0, condition_data0, False)
       If combine_line_value_with_dpoint_pair_ > 1 Then
        Exit Function
       End If
    combine_line_value_with_dpoint_pair_ = set_general_string( _
         it%, 0, 0, 0, "1", "0", "0", "0", time_string(line_value(l%).data(0).data0.value, _
            line_value(tn%).data(0).data0.value, True, False), 0, 0, 1, temp_record, 0, 0)
         Call set_level_(Ddpoint_pair(dp%).record_.no_reduce, 4)
       If combine_line_value_with_dpoint_pair_ > 1 Then
        Exit Function
       End If
     Else
       temp_record.record_data = re
        Call add_conditions_to_record(line_value_, tn%, 0, 0, temp_record.record_data.data0.condition_data)
         combine_line_value_with_dpoint_pair_ = set_general_string_from_relation( _
          0, 0, Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%), _
           Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1), _
            Ddpoint_pair(dp%).data(0).data0.poi(2 * n3), _
             Ddpoint_pair(dp%).data(0).data0.poi(2 * n3 + 1), _
              Ddpoint_pair(dp%).data(0).data0.poi(2 * n4%), _
               Ddpoint_pair(dp%).data(0).data0.poi(2 * n4% + 1), _
          0, 0, Ddpoint_pair(dp%).data(0).data0.n(2 * n2%), _
           Ddpoint_pair(dp%).data(0).data0.n(2 * n2% + 1), _
            Ddpoint_pair(dp%).data(0).data0.n(2 * n3), _
             Ddpoint_pair(dp%).data(0).data0.n(2 * n3 + 1), _
              Ddpoint_pair(dp%).data(0).data0.n(2 * n4%), _
               Ddpoint_pair(dp%).data(0).data0.n(2 * n4% + 1), _
          0, Ddpoint_pair(dp%).data(0).data0.line_no(n2%), _
           Ddpoint_pair(dp%).data(0).data0.line_no(n3%), _
            Ddpoint_pair(dp%).data(0).data0.line_no(n4%), _
                line_value(l%).data(0).data0.value, "1", temp_record, 0)
         Call set_level_(Ddpoint_pair(dp%).record_.no_reduce, 4)
        If combine_line_value_with_dpoint_pair_ > 1 Then
         Exit Function
        End If
     End If
End Function

Public Function combine_relation_with_dpoint_pair00_(ByVal dp%, _
            ByVal n1%, ByVal n2%, ByVal n3%, ByVal n4%, _
              ByVal v$, re As record_data_type) As Byte
Dim temp_record As total_record_type
Dim tn%
Dim tv$
Dim m(1) As Integer
If n1% < 4 Or (n1% And Ddpoint_pair(dp%).data(0).data0.con_line_type(0) = _
                Ddpoint_pair(dp%).data(0).data0.con_line_type(1)) Then
  If is_line_value(Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%), _
        Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1), _
          Ddpoint_pair(dp%).data(0).data0.n(2 * n3%), _
           Ddpoint_pair(dp%).data(0).data0.n(2 * n3% + 1), _
            Ddpoint_pair(dp%).data(0).data0.line_no(n3%), _
             "", tn%, -1000, 0, 0, 0, _
              line_value_data0) = 1 Then
  temp_record.record_data = re
   Call add_conditions_to_record(line_value_, tn%, 0, 0, temp_record.record_data.data0.condition_data)
    combine_relation_with_dpoint_pair00_ = set_line_value( _
     Ddpoint_pair(dp%).data(0).data0.poi(2 * n4%), _
      Ddpoint_pair(dp%).data(0).data0.poi(2 * n4% + 1), _
       divide_string(line_value(tn%).data(0).data0.value, v$, True, False), _
        0, 0, 0, temp_record, 0, 0, False)
   If combine_relation_with_dpoint_pair00_ > 1 Then
    Exit Function
   End If
  ElseIf is_line_value(Ddpoint_pair(dp%).data(0).data0.poi(2 * n4%), _
        Ddpoint_pair(dp%).data(0).data0.poi(2 * n4% + 1), _
          Ddpoint_pair(dp%).data(0).data0.n(2 * n4%), _
           Ddpoint_pair(dp%).data(0).data0.n(2 * n4% + 1), _
            Ddpoint_pair(dp%).data(0).data0.line_no(n4%), _
             "", tn%, -1000, 0, 0, 0, _
              line_value_data0) = 1 Then
  temp_record.record_data = re
   Call add_conditions_to_record(line_value_, tn%, 0, 0, temp_record.record_data.data0.condition_data)
    combine_relation_with_dpoint_pair00_ = set_line_value( _
     Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%), _
      Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1), _
       time_string(line_value(tn%).data(0).data0.value, v$, True, False), _
        0, 0, 0, temp_record, 0, 0, False)
   If combine_relation_with_dpoint_pair00_ > 1 Then
    Exit Function
   End If
  Else
    temp_record.record_data = re
   combine_relation_with_dpoint_pair00_ = set_Drelation( _
    Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%), _
     Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1), _
      Ddpoint_pair(dp%).data(0).data0.poi(2 * n4%), _
       Ddpoint_pair(dp%).data(0).data0.poi(2 * n4% + 1), _
    Ddpoint_pair(dp%).data(0).data0.n(2 * n3%), _
     Ddpoint_pair(dp%).data(0).data0.n(2 * n3% + 1), _
      Ddpoint_pair(dp%).data(0).data0.n(2 * n4%), _
       Ddpoint_pair(dp%).data(0).data0.n(2 * n4% + 1), _
        Ddpoint_pair(dp%).data(0).data0.line_no(n3%), _
         Ddpoint_pair(dp%).data(0).data0.line_no(n4%), _
          v$, temp_record, 0, 0, 0, 0, 0, False)
   If combine_relation_with_dpoint_pair00_ > 1 Then
    Exit Function
   End If
  End If
ElseIf n1% = 4 Then
 If n2% = 0 Then
 m(0) = 3
 m(1) = 2
 ElseIf n2% = 1 Then
 m(0) = 2
 m(1) = 3
 End If
 If Ddpoint_pair(dp%).data(0).data0.con_line_type(0) = 3 Or _
     Ddpoint_pair(dp%).data(0).data0.con_line_type(0) = 5 Then
   tv$ = minus_string(v$, "1", True, False)
 ElseIf Ddpoint_pair(dp%).data(0).data0.con_line_type(0) = 4 Or _
     Ddpoint_pair(dp%).data(0).data0.con_line_type(0) = 8 Then
  If n1% = 0 Then
   tv$ = minus_string("1", v$, True, False)
  ElseIf n1% = 1 Then
   tv$ = add_string("1", v$, True, False)
  End If
 ElseIf Ddpoint_pair(dp%).data(0).data0.con_line_type(0) = 6 Or _
     Ddpoint_pair(dp%).data(0).data0.con_line_type(0) = 7 Then
  If n1% = 0 Then
   tv$ = add_string(v$, "1", True, False)
  ElseIf n1% = 1 Then
   tv$ = minus_string("1", v$, True, False)
  End If
 End If
 temp_record.record_data = re
   combine_relation_with_dpoint_pair00_ = set_Drelation( _
    Ddpoint_pair(dp%).data(0).data0.poi(2 * m(0)), _
     Ddpoint_pair(dp%).data(0).data0.poi(2 * m(0) + 1), _
      Ddpoint_pair(dp%).data(0).data0.poi(2 * m(1)), _
       Ddpoint_pair(dp%).data(0).data0.poi(2 * m(1) + 1), _
    Ddpoint_pair(dp%).data(0).data0.n(2 * m(0)), _
     Ddpoint_pair(dp%).data(0).data0.n(2 * m(0) + 1), _
      Ddpoint_pair(dp%).data(0).data0.n(2 * m(1)), _
       Ddpoint_pair(dp%).data(0).data0.n(2 * m(1) + 1), _
        Ddpoint_pair(dp%).data(0).data0.line_no(m(0)), _
         Ddpoint_pair(dp%).data(0).data0.line_no(m(1)), _
          tv$, temp_record, 0, 0, 0, 0, 0, False)
   If combine_relation_with_dpoint_pair00_ > 1 Then
    Exit Function
   End If
ElseIf n1% = 5 Then
 If n2% = 2 Then
 m(0) = 1
 m(1) = 0
 ElseIf n2% = 2 Then
 m(0) = 0
 m(1) = 1
 End If
temp_record.record_data = re
 If Ddpoint_pair(dp%).data(0).data0.con_line_type(1) = 3 Or _
     Ddpoint_pair(dp%).data(0).data0.con_line_type(1) = 5 Then
   tv$ = minus_string(v$, "1", True, False)
 ElseIf Ddpoint_pair(dp%).data(0).data0.con_line_type(1) = 4 Or _
     Ddpoint_pair(dp%).data(0).data0.con_line_type(1) = 6 Then
  If n1% = 2 Then
   tv$ = minus_string("1", v$, True, False)
  ElseIf n1% = 3 Then
   tv$ = add_string("1", v$, True, False)
  End If
 ElseIf Ddpoint_pair(dp%).data(0).data0.con_line_type(1) = 7 Or _
     Ddpoint_pair(dp%).data(0).data0.con_line_type(1) = 8 Then
  If n1% = 2 Then
   tv$ = add_string(v$, "1", True, False)
  ElseIf n1% = 3 Then
   tv$ = minus_string("1", v$, True, False)
  End If
 End If
 temp_record.record_data = re
   combine_relation_with_dpoint_pair00_ = set_Drelation( _
    Ddpoint_pair(dp%).data(0).data0.poi(2 * m(0)), _
     Ddpoint_pair(dp%).data(0).data0.poi(2 * m(0) + 1), _
      Ddpoint_pair(dp%).data(0).data0.poi(2 * m(1)), _
       Ddpoint_pair(dp%).data(0).data0.poi(2 * m(1) + 1), _
    Ddpoint_pair(dp%).data(0).data0.n(2 * m(0)), _
     Ddpoint_pair(dp%).data(0).data0.n(2 * m(0) + 1), _
      Ddpoint_pair(dp%).data(0).data0.n(2 * m(1)), _
       Ddpoint_pair(dp%).data(0).data0.n(2 * m(1) + 1), _
        Ddpoint_pair(dp%).data(0).data0.line_no(m(0)), _
         Ddpoint_pair(dp%).data(0).data0.line_no(m(1)), _
          tv$, temp_record, 0, 0, 0, 0, 0, False)
   If combine_relation_with_dpoint_pair00_ > 1 Then
    Exit Function
   End If
End If
End Function
Public Function combine_relation_with_dpoint_pair01_(ByVal dp%, _
             ByVal n1%, ByVal n2%, ByVal n3%, ByVal n4%, _
               p() As Integer, n() As Integer, l() As Integer, v() As String, _
                ByVal para As String, ty As Byte, re As record_data_type, no_reduce As Byte) As Byte
Dim temp_record As total_record_type
Dim it(1) As Integer
Dim tp(7) As Integer
Dim tn(7) As Integer
Dim tl(1) As Integer
Dim Tpara(3) As String
Dim ty_(2)  As Boolean
Dim l_p() As Integer
Dim temp_record0 As record_type0
temp_record.record_data = re
Tpara(0) = para
If ty = 0 Then
tp(0) = p(0)
tp(1) = p(1)
tn(0) = n(0)
tn(1) = n(1)
tl(0) = l(0)
ElseIf ty = 1 Then
tp(0) = p(2)
tp(1) = p(3)
tn(0) = n(2)
tn(1) = n(3)
tl(0) = l(1)
ElseIf ty = 2 Then
tp(0) = p(4)
tp(1) = p(5)
tn(0) = n(4)
tn(1) = n(5)
tl(0) = l(2)
End If
ty_(0) = True

l_p(0) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%)
l_p(1) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1)
ty_(1) = combine_relation_with_line_(p(), v(), l_p(), Tpara(1), ty)
If ty_(1) Then
tp(2) = tp(0)
tp(3) = tp(1)
tn(2) = tn(0)
tn(3) = tn(1)
tl(1) = tl(0)
Else
tp(2) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n2%)
tp(3) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n2% + 1)
tn(2) = Ddpoint_pair(dp%).data(0).data0.n(2 * n2%)
tn(3) = Ddpoint_pair(dp%).data(0).data0.n(2 * n2% + 1)
tl(1) = Ddpoint_pair(dp%).data(0).data0.line_no(n2%)
Tpara(1) = "1"
End If
'***
l_p(0) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%)
l_p(1) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1)
ty_(2) = combine_relation_with_line_(p(), v(), l_p(), Tpara(2), ty)
If ty_(2) Then
tp(4) = tp(0)
tp(5) = tp(1)
tn(4) = tn(0)
tn(5) = tn(1)
tl(2) = tl(0)
Else
tp(4) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n3%)
tp(5) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n3% + 1)
tn(4) = Ddpoint_pair(dp%).data(0).data0.n(2 * n3%)
tn(5) = Ddpoint_pair(dp%).data(0).data0.n(2 * n3% + 1)
tl(2) = Ddpoint_pair(dp%).data(0).data0.line_no(n3%)
Tpara(2) = "1"
End If
'**
l_p(0) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n4%)
l_p(1) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n4% + 1)
ty_(3) = combine_relation_with_line_(p(), v(), l_p(), Tpara(3), ty)
If ty_(3) Then
tp(6) = tp(0)
tp(7) = tp(1)
tn(6) = tn(0)
tn(7) = tn(1)
tl(3) = tl(0)
Else
tp(6) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n4%)
tp(7) = Ddpoint_pair(dp%).data(0).data0.poi(2 * n4% + 1)
tn(6) = Ddpoint_pair(dp%).data(0).data0.n(2 * n4%)
tn(7) = Ddpoint_pair(dp%).data(0).data0.n(2 * n4% + 1)
tl(3) = Ddpoint_pair(dp%).data(0).data0.line_no(n4%)
Tpara(3) = "1"
End If
If ty_(1) Or ty_(2) Or ty_(3) Then
 If divide_string(Tpara(0), Tpara(1), True, False) = divide_string(Tpara(2), Tpara(3), True, False) Then
  combine_relation_with_dpoint_pair01_ = set_dpoint_pair(tp(0), tp(1), tp(2), tp(3), _
    tp(4), tp(5), tp(6), tp(7), tn(0), tn(1), tn(2), tn(3), tn(4), tn(5), _
     tn(6), tn(7), tl(0), tl(1), tl(2), tl(3), 0, temp_record, True, 0, 0, 0, 0, False)
 Else
  combine_relation_with_dpoint_pair01_ = _
     set_item0(tp(0), tp(1), tp(6), tp(7), "*", tn(0), tn(1), tn(6), tn(7), _
        tl(0), tl(3), "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, _
          0, it(0), no_reduce, 0, condition_data0, False)
  If combine_relation_with_dpoint_pair01_ > 1 Then
   Exit Function
  End If
  combine_relation_with_dpoint_pair01_ = _
     set_item0(tp(2), tp(3), tp(4), tp(5), "*", tn(2), tn(3), tn(4), tn(5), _
        tl(1), tl(2), "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, _
           0, it(0), no_reduce, 0, condition_data0, False)
  If combine_relation_with_dpoint_pair01_ > 1 Then
   Exit Function
  End If
  combine_relation_with_dpoint_pair01_ = set_general_string(it(0), it(1), 0, 0, _
     time_string(Tpara(0), Tpara(3), True, False), _
      time_string("-1", time_string(Tpara(1), Tpara(2), False, False), True, False), _
      "0", "0", "0", 0, 0, 0, temp_record, 0, 0)
 End If
End If
End Function
Public Function combine_line_value_with_line_value(ByVal l_v%, _
        no_reduce As Byte) As Byte
Dim i%, j%, k%, no%
Dim last_tn%
Dim tn() As Integer
Dim temp_record As total_record_type
Dim m(1) As Integer
Dim n(1) As Integer
Dim n_(1) As Integer
Dim lv As line_value_data0_type
For i% = 0 To 1
n(0) = i%
n(1) = (i% + 1) Mod 2
For j% = 0 To 1
m(0) = j%
m(1) = (j% + 1) Mod 2
lv.poi(m(0)) = line_value(l_v%).data(0).data0.poi(n(0))
lv.poi(m(1)) = -1
Call search_for_line_value(lv, m(0), n_(0), 1)
lv.poi(m(1)) = 30000
Call search_for_line_value(lv, m(0), n_(1), 1)  '5.7
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
 no% = line_value(k%).data(0).record.data1.index.i(m(0))
 If no% < l_v% And no% > 0 Then
  If line_value(no%).record_.no_reduce < 255 Then
 last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
 End If
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
Call add_conditions_to_record(line_value_, l_v%, no%, 0, temp_record.record_data.data0.condition_data)
'combine_line_value_with_line_value = set_property_of_relation( _
      line_value(l_v%).data(0).data0.poi(0), line_value(l_v%).data(0).data0.poi(1), _
       line_value(no%).data(0).data0.poi(0), line_value(no%).data(0).data0.poi(1), _
         divide_string(line_value(l_v%).data(0).data0.value, _
          line_value(no%).data(0).data0.value, True, False), line_value(l_v%).data(0).data0.n(0), _
           line_value(l_v%).data(0).data0.n(1), line_value(no%).data(0).data0.n(0), _
            line_value(no%).data(0).data0.n(1), line_value(l_v%).data(0).data0.line_no, _
             line_value(no%).data(0).data0.line_no, temp_record, no_reduce)
     temp_record.record_data.data0.theorem_no = 1
      'temp_record.no_reduce = 0
 If line_value(l_v).data(0).data0.line_no = line_value(no%).data(0).data0.line_no Then
  If m(0) = 0 And n(0) = 0 Then
    If line_value(l_v%).data(0).data0.n(1) > line_value(no%).data(0).data0.n(1) Then
     combine_line_value_with_line_value = set_line_value( _
      line_value(no%).data(0).data0.poi(m(1)), line_value(l_v%).data(0).data0.poi(n(1)), _
       minus_string(line_value(l_v%).data(0).data0.value, line_value(no%).data(0).data0.value, True, False), _
        line_value(no%).data(0).data0.n(m(1)), line_value(l_v%).data(0).data0.n(n(1)), _
         line_value(l_v%).data(0).data0.line_no, temp_record, 0, no_reduce, False)
     If combine_line_value_with_line_value > 1 Then
      Exit Function
     End If
   ElseIf line_value(l_v%).data(0).data0.n(1) < line_value(no%).data(0).data0.n(1) Then
     combine_line_value_with_line_value = set_line_value( _
      line_value(no%).data(0).data0.poi(m(1)), line_value(l_v%).data(0).data0.poi(n(1)), _
       minus_string(line_value(no%).data(0).data0.value, line_value(l_v%).data(0).data0.value, True, False), _
        line_value(no%).data(0).data0.n(m(1)), line_value(l_v%).data(0).data0.n(n(1)), _
         line_value(l_v%).data(0).data0.line_no, temp_record, 0, no_reduce, False)
    If combine_line_value_with_line_value > 1 Then
     Exit Function
    End If
   End If
  ElseIf (m(0) = 0 And n(0) = 1) Or (m(0) = 1 And n(0) = 0) Then
    combine_line_value_with_line_value = set_line_value( _
     line_value(no%).data(0).data0.poi(m(1)), line_value(l_v%).data(0).data0.poi(n(1)), _
      add_string(line_value(l_v%).data(0).data0.value, line_value(no%).data(0).data0.value, True, False), _
       line_value(no%).data(0).data0.n(m(1)), line_value(l_v%).data(0).data0.n(n(1)), _
        line_value(l_v%).data(0).data0.line_no, temp_record, 0, no_reduce, False)
     If combine_line_value_with_line_value > 1 Then
      Exit Function
     End If
  ElseIf m(0) = 1 And n(0) = 1 Then
   If line_value(l_v%).data(0).data0.n(0) > line_value(no%).data(0).data0.n(0) Then
      combine_line_value_with_line_value = set_line_value( _
       line_value(no%).data(0).data0.poi(m(1)), line_value(l_v%).data(0).data0.poi(n(1)), _
        minus_string(line_value(no%).data(0).data0.value, line_value(l_v%).data(0).data0.value, True, False), _
         line_value(no%).data(0).data0.n(m(1)), line_value(l_v%).data(0).data0.n(n(1)), _
          line_value(l_v%).data(0).data0.line_no, temp_record, 0, no_reduce, False)
     If combine_line_value_with_line_value > 1 Then
      Exit Function
     End If
   ElseIf line_value(l_v%).data(0).data0.n(0) < line_value(no%).data(0).data0.n(0) Then
     combine_line_value_with_line_value = set_line_value( _
      line_value(no%).data(0).data0.poi(m(1)), line_value(l_v%).data(0).data0.poi(n(1)), _
       minus_string(line_value(l_v%).data(0).data0.value, line_value(no%).data(0).data0.value, True, False), _
        line_value(no%).data(0).data0.n(m(1)), line_value(l_v%).data(0).data0.n(n(1)), _
         line_value(l_v%).data(0).data0.line_no, temp_record, 0, no_reduce, False)
     If combine_line_value_with_line_value > 1 Then
      Exit Function
     End If
   End If
  End If
 End If
Next k%
Next j%
Next i%
End Function


Public Sub init_reduce()
reduce_level = reduce_level0
'**********
End Sub
Public Function combine_two_polygon(po1 As polygon, po2 As polygon, opo1 As polygon, opo2 As polygon, _
                     relation As relation_data0_type) As Integer
'1=+.-1=-,0
Dim i%, j%, ty%
Dim tp(3) As Integer
Dim tp1(3) As Integer
Dim tp2(3) As Integer
Dim tp3(3) As Integer
Dim total_v(1) As Integer
total_v(0) = po1.total_v
 total_v(1) = po2.total_v
For i% = 0 To 3
 tp(i%) = po1.v(i%)
  tp1(i%) = po2.v(i%)
Next i%
For i% = 0 To total_v(0) - 1
 For j% = 0 To total_v(1) - 1
  If tp(i%) = tp1(j%) Then '一点同
   If tp((i% + 1) Mod total_v(0)) = tp1((j% + 1) Mod total_v(1)) Then
     tp2(0) = tp(i%)
     tp2(1) = tp((i% + 1) Mod total_v(0))
     tp2(2) = tp((i% + 2) Mod total_v(0))
     tp2(3) = tp((i% + 3) Mod total_v(0))
     tp3(0) = tp1(j%)
     tp3(1) = tp1((j% + 1) Mod total_v(1))
     tp3(2) = tp1((j% + 2) Mod total_v(1))
     tp3(3) = tp1((j% + 3) Mod total_v(1))
     GoTo combine_two_polygon_mark1
   ElseIf tp((i% + 1) Mod total_v(0)) = tp1((j% + total_v(1) - 1) Mod total_v(1)) Then
     tp2(0) = tp(i%)
     tp2(1) = tp((i% + 1) Mod total_v(0))
     tp2(2) = tp((i% + 2) Mod total_v(0))
     tp2(3) = tp((i% + 3) Mod total_v(0))
     tp3(0) = tp1(j%)
     tp3(1) = tp1((j% + total_v(1) - 1) Mod total_v(1))
     tp3(2) = tp1((j% + total_v(1) - 2) Mod total_v(1))
     tp3(3) = tp1((j% + total_v(1) - 3) Mod total_v(1))
     GoTo combine_two_polygon_mark1
   ElseIf tp((i% + total_v(0) - 1) Mod total_v(0)) = tp1((j% + 1) Mod total_v(1)) Then
     tp2(0) = tp(i%)
     tp2(1) = tp((i% + total_v(0) - 1) Mod total_v(0))
     tp2(2) = tp((i% + total_v(0) - 2) Mod total_v(0))
     tp2(3) = tp((i% + total_v(0) - 3) Mod total_v(0))
     tp3(0) = tp1(j%)
     tp3(1) = tp1((j% + 1) Mod total_v(1))
     tp3(2) = tp1((j% + 2) Mod total_v(1))
     tp3(3) = tp1((j% + 3) Mod total_v(1))
     GoTo combine_two_polygon_mark1
   ElseIf tp((i% + total_v(0) - 1) Mod total_v(0)) = tp1((j% + total_v(1) - 1) Mod total_v(1)) Then
     tp2(0) = tp(i%)
     tp2(1) = tp((i% + total_v(0) - 1) Mod total_v(0))
     tp2(2) = tp((i% + total_v(0) - 2) Mod total_v(0))
     tp2(3) = tp((i% + total_v(0) - 3) Mod total_v(0))
     tp3(0) = tp1(j%)
     tp3(1) = tp1((j% + total_v(1) - 1) Mod total_v(1))
     tp3(2) = tp1((j% + total_v(1) - 2) Mod total_v(1))
     tp3(3) = tp1((j% + total_v(1) - 3) Mod total_v(1))
     GoTo combine_two_polygon_mark1
   ElseIf total_v(0) = 3 And total_v(1) = 3 Then
     tp2(0) = tp(i%)
     tp2(1) = tp((i% + 1) Mod total_v(0))
     tp2(2) = tp((i% + 2) Mod total_v(0))
     tp2(3) = tp((i% + 3) Mod total_v(0))
     tp3(0) = tp1(j%)
     tp3(1) = tp1((j% + 1) Mod total_v(1))
     tp3(2) = tp1((j% + 2) Mod total_v(1))
     tp3(3) = tp1((j% + 3) Mod total_v(1))
     GoTo combine_two_polygon_mark1
   End If 'Else
  End If
  Next j%
  Next i%
  Exit Function
 '两个多边形tp2(0)=tp3(0),tp2(1)=tp3(1)
 '两点同
combine_two_polygon_mark1:
 If total_v(0) = 3 Then
  tp2(3) = tp2(2) '三角形
 End If
 If total_v(1) = 3 Then
  tp3(3) = tp3(2)
 End If
 combine_two_polygon = combine_two_polygon0(tp2(3), tp2(0), tp2(1), tp2(2), total_v(0), _
    tp3(3), tp3(0), tp3(1), tp3(2), total_v(1), opo1, opo2, relation)
combine_two_polygon_mark0:
 End Function

Public Function combine_two_polygon0(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal v1%, _
         ByVal p5%, ByVal p6%, ByVal p7%, ByVal p8%, ByVal v2%, po1 As polygon, po2 As polygon, _
          relation As relation_data0_type) As Integer
    'p2%=p6%,p3%=p7%'1 v1+v2=po.-1 v1-v2=po ,-2 v2-v1=po,-3 v1-v2%=po1+po2,-4,v2-v1=po1+po2,
    '2,v1+v2%=po1-po2,3  ,v1+v2%=po2+po1 -5 v1-v2=等价po1-po2
    'case -5   v1-v2=po3-po4
Dim tn1(3) As Integer
Dim tn2(3) As Integer
Dim tl(3) As Byte
Dim tp(3) As Integer
Dim tA(5) As Integer
Dim out%, t_p%
Dim tpol As polygon
Dim v%
Dim rela As relation_data0_type
Dim cond As condition_data_type
Dim temp_record As total_record_type
  tl(0) = line_number0(p2%, p1%, tn1(0), tn1(1))  '邻边
    tl(1) = line_number0(p6%, p5%, tn1(2), tn1(3))
  tl(2) = line_number0(p3%, p4%, tn2(0), tn2(1))  '邻边
    tl(3) = line_number0(p7%, p8%, tn2(2), tn2(3))
If v1% = 3 And v2% = 3 Then '两个三角形
  If p3% = p7% Then '两点同
    If tl(0) = tl(1) Then '邻边同
      If (tn1(0) < tn1(1) And tn1(2) < tn1(3) And tn1(3) < tn1(1)) Or _
           (tn1(0) > tn1(1) And tn1(2) > tn1(3) And tn1(3) > tn1(1)) Then
             po1.total_v = 3
              po1.v(0) = p5%
               po1.v(1) = p3%
                po1.v(2) = p1%
           combine_two_polygon0 = -1
            Call arrange_four_point(p2%, p1%, p2%, p5%, tn1(0), tn1(1), tn1(0), tn1(3), tl(0), tl(0), _
                relation.poi(0), relation.poi(1), relation.poi(2), relation.poi(3), 0, 0, _
                 relation.n(0), relation.n(1), relation.n(2), relation.n(3), 0, 0, relation.line_no(0), _
                  relation.line_no(1), 0, relation.ty, cond, 0)
            Exit Function
      ElseIf (tn1(0) > tn1(1) And tn1(2) > tn1(3) And tn1(3) < tn1(1)) Or _
             (tn1(0) < tn1(1) And tn1(2) < tn1(3) And tn1(3) > tn1(1)) Then
             po1.total_v = 3
              po1.v(0) = p1%
               po1.v(1) = p5%
                po1.v(2) = p3%
                 combine_two_polygon0 = -2
             Call arrange_four_point(p2%, p1%, p2%, p5%, tn1(0), tn1(1), tn1(0), tn1(3), tl(0), tl(0), _
                relation.poi(0), relation.poi(1), relation.poi(2), relation.poi(3), 0, 0, _
                 relation.n(0), relation.n(1), relation.n(2), relation.n(3), 0, 0, relation.line_no(0), _
                  relation.line_no(1), 0, relation.ty, cond, 0)
               Exit Function
      ElseIf (tn1(0) < tn1(1) And tn1(2) > tn1(3)) Or _
              (tn1(0) > tn1(1) And tn1(2) < tn1(3)) Then
            po1.total_v = 3
             po1.v(0) = p1%
               po1.v(1) = p5%
                 po1.v(2) = p3%
            Call arrange_four_point(p2%, p1%, p2%, p5%, tn1(0), tn1(1), tn1(0), tn1(3), tl(0), tl(0), _
                relation.poi(0), relation.poi(1), relation.poi(2), relation.poi(3), 0, 0, _
                 relation.n(0), relation.n(1), relation.n(2), relation.n(3), 0, 0, relation.line_no(0), _
                  relation.line_no(1), 0, relation.ty, cond, 0)
               combine_two_polygon0 = 1
                 Exit Function
      End If
    ElseIf tl(2) = tl(3) Then
      If (tn2(0) < tn2(1) And tn2(2) < tn2(3) And tn2(3) < tn2(1)) Or _
            (tn2(0) > tn2(1) And tn2(2) > tn2(2) And tn2(3) > tn2(1)) Then
            po1.total_v = 3
             po1.v(0) = p5%
              po1.v(1) = p1%
               po1.v(2) = p2%
             Call arrange_four_point(p3%, p1%, p3%, p5%, tn2(0), tn2(1), tn2(0), tn2(3), tl(2), tl(2), _
                relation.poi(0), relation.poi(1), relation.poi(2), relation.poi(3), 0, 0, _
                 relation.n(0), relation.n(1), relation.n(2), relation.n(3), 0, 0, relation.line_no(0), _
                  relation.line_no(1), 0, relation.ty, cond, 0)
               combine_two_polygon0 = -1
                Exit Function
      ElseIf (tn2(0) > tn2(1) And tn2(2) > tn2(3) And tn2(3) < tn2(1)) Or _
          (tn2(0) < tn2(1) And tn2(2) < tn2(3) And tn2(3) > tn2(1)) Then
            po1.total_v = 3
             po1.v(0) = p5%
              po1.v(1) = p1%
                po1.v(2) = p2%
                 combine_two_polygon0 = -2
             Call arrange_four_point(p3%, p1%, p3%, p5%, tn2(0), tn2(1), tn2(0), tn2(3), tl(2), tl(2), _
                relation.poi(0), relation.poi(1), relation.poi(2), relation.poi(3), 0, 0, _
                 relation.n(0), relation.n(1), relation.n(2), relation.n(3), 0, 0, relation.line_no(0), _
                  relation.line_no(1), 0, relation.ty, cond, 0)
                   Exit Function
      ElseIf (tn2(0) < tn2(1) And tn2(2) > tn2(3)) Or _
                   (tn2(0) > tn2(1) And tn2(2) < tn2(3)) Then
             po1.total_v = 3
              po1.v(0) = p5%
               po1.v(1) = p1%
                po1.v(2) = p2%
                 combine_two_polygon0 = 1
              Call arrange_four_point(p3%, p1%, p3%, p5%, tn2(0), tn2(1), tn2(0), tn2(3), tl(2), tl(2), _
                relation.poi(0), relation.poi(1), relation.poi(2), relation.poi(3), 0, 0, _
                 relation.n(0), relation.n(1), relation.n(2), relation.n(3), 0, 0, relation.line_no(0), _
                  relation.line_no(1), 0, relation.ty, cond, 0)
                 Exit Function
      End If
    Else
      tA(0) = angle_number(p2%, p1%, p3%, "", 0)
      tA(1) = angle_number(p6%, p5%, p7%, "", 0)
       If (tA(0) > 0 And tA(1) < 0) Or (tA(0) < 0 And tA(1) > 0) Then '两个三角形合并为一个四边形
        tA(0) = angle_number(p1%, p2%, p3%, "", 0)
        tA(1) = angle_number(p1%, p2%, p5%, "", 0)
        tA(2) = angle_number(p1%, p3%, p2%, "", 0)
        tA(3) = angle_number(p1%, p3%, p5%, "", 0)
        'tA(4) = angle_number(p3%, p2%, p5%, "", 0) '大角
        'tA(5) = angle_number(p2%, p3%, p5%, "", 0)
        If (tA(0) > 0 And tA(1) > 0) Or (tA(0) < 0 And tA(1) < 0) Then
            tA(0) = 1 '和角<180
        Else
            tA(0) = -1
        End If
        If (tA(2) > 0 And tA(3) > 0) Or (tA(2) < 0 And tA(3) < 0) Then
            tA(2) = 1 '和角<180
        Else
            tA(2) = -1
        End If
        If tA(0) > 0 And tA(2) > 0 Then '四边形
             po1.total_v = 4
              po1.v(0) = p1%
               po1.v(1) = p2%
                po1.v(2) = p5%
                 po1.v(3) = p3%
                  combine_two_polygon0 = 1
                   Exit Function
        ElseIf tA(0) > 0 And tA(2) < 0 Then
             po1.total_v = 3
              po1.v(0) = p1%
               po1.v(1) = p2%
                po1.v(2) = p5%
             po2.total_v = 3
              po2.v(0) = p1%
               po2.v(1) = p3%
                po2.v(2) = p5%
                 combine_two_polygon0 = 2
                   Exit Function
        ElseIf tA(0) < 0 And tA(2) > 0 Then
            po1.total_v = 3
              po1.v(0) = p1%
               po1.v(1) = p3%
                po1.v(2) = p5%
             po2.total_v = 3
              po2.v(0) = p1%
               po2.v(1) = p2%
                po2.v(2) = p5%
                 combine_two_polygon0 = 2
                   Exit Function
        End If
      Else '同一边
       tA(0) = angle_number(p1%, p2%, p3%, "", 0)
       tA(1) = angle_number(p1%, p2%, p5%, "", 0)
       tA(2) = angle_number(p1%, p3%, p2%, "", 0)
       tA(3) = angle_number(p1%, p3%, p5%, "", 0)
        If (tA(0) > 0 And tA(1) > 0) Or (tA(0) < 0 And tA(1)) < 0 Then
            tA(0) = 1
        Else
            tA(0) = -1
        End If
        If (tA(2) > 0 And tA(3) > 0) Or (tA(2) < 0 And tA(3)) < 0 Then
            tA(2) = 1
        Else
            tA(2) = -1
        End If
       If tA(0) > 0 And tA(2) > 0 Then
            po1.total_v = 3
              po1.v(0) = p1%
               po1.v(1) = p2%
                po1.v(2) = p5%
             po2.total_v = 3
              po2.v(0) = p1%
               po2.v(1) = p3%
                po2.v(2) = p5%
                 combine_two_polygon0 = -3
                   Exit Function
       ElseIf tA(0) < 0 And tA(2) < 0 Then
            po1.total_v = 3
              po1.v(0) = p1%
               po1.v(1) = p2%
                po1.v(2) = p5%
             po2.total_v = 3
              po2.v(0) = p1%
               po2.v(1) = p3%
                po2.v(2) = p5%
                 combine_two_polygon0 = -4
                   Exit Function
       ElseIf tA(0) > 0 And tA(1) < 0 Then
         tp(0) = is_line_line_intersect( _
              line_number0(p1%, p3%, 0, 0), line_number0(p2%, p5%, 0, 0), 0, 0, False)
          If tp(0) > 0 Then
           po1.total_v = 3
            po1.v(0) = p1%
             po1.v(1) = p2%
              po1.v(2) = tp(0)
           po2.total_v = 3
            po2.v(0) = p3%
             po2.v(1) = p5%
              po2.v(2) = tp(0)
                 combine_two_polygon0 = -5
          Else
                 combine_two_polygon0 = 0
          End If
                   Exit Function
       ElseIf tA(0) < 0 And tA(1) > 0 Then
          tp(0) = is_line_line_intersect( _
              line_number0(p1%, p2%, 0, 0), line_number0(p3%, p5%, 0, 0), 0, 0, False)
          If tp(0) > 0 Then
           po1.total_v = 3
            po1.v(0) = p1%
             po1.v(1) = p3%
              po1.v(2) = tp(0)
           po2.total_v = 3
            po2.v(0) = p5%
             po2.v(1) = p6%
              po2.v(2) = tp(0)
                 combine_two_polygon0 = -5
          Else
                 combine_two_polygon0 = 0
          End If
                   Exit Function
       End If
      End If
     End If
    Else '一点同
      tl(2) = line_number0(p2%, p3%, tn2(0), tn2(1))  '邻边
      tl(3) = line_number0(p6%, p7%, tn2(2), tn2(3))
      If tl(0) = tl(1) And tl(2) = tl(3) Then
       combine_two_polygon0 = combine_two_triangle_with_one_co_point( _
        tl(0), tn1(0), tn1(1), tn1(3), tl(2), tn2(0), tn2(1), tn2(3), po1, po2)
      ElseIf tl(0) = tl(3) And tl(1) = tl(2) Then
       combine_two_polygon0 = combine_two_triangle_with_one_co_point( _
        tl(0), tn1(0), tn1(1), tn2(3), tl(2), tn2(0), tn2(1), tn1(3), po1, po2)
      End If
    End If
  ElseIf v1% = 3 And v2% = 4 Then
    If p1% = p5% Then '分为两个三角形
             po1.total_v = 3
              po1.v(0) = p7%
               po1.v(1) = p8%
                po1.v(2) = p5%
                  combine_two_polygon0 = -2
                   Exit Function
   ElseIf p1% = p8% Then '
            po1.total_v = 3
              po1.v(0) = p8%
               po1.v(1) = p5%
                po1.v(2) = p6%
                  combine_two_polygon0 = -2
                   Exit Function
   ElseIf tl(0) = tl(1) And tl(2) = tl(3) Then
           '四边形含在三角形中
            po1.total_v = 3
              po1.v(0) = p1%
               po1.v(1) = p5%
                po1.v(2) = p8%
      If (tn1(0) > tn1(1) And tn1(0) > tn1(3)) Or _
           (tn1(0) < tn1(1) And tn1(0) < tn1(3)) Then
                  combine_two_polygon0 = -1
      Else
                  combine_two_polygon0 = 1
      End If
                   Exit Function
   ElseIf tl(0) = tl(1) Then
     combine_two_polygon0 = combine_triangle_with_polygon(p1%, p2%, p3%, p5%, p6%, p7%, p8%, _
        tn1(0), tn1(1), tn1(2), tn1(3), po1, po2)
   ElseIf tl(2) = tl(3) Then
     combine_two_polygon0 = combine_triangle_with_polygon(p1%, p3%, p2%, p8%, p7%, p6%, p5%, _
        tn2(0), tn2(1), tn2(2), tn2(3), po1, po2)
   End If
  ElseIf v1% = 4 And v2% = 3 Then
   out% = combine_two_polygon0(p5%, p6%, p7%, p8%, v2%, _
         p1%, p2%, p3%, p4%, ByVal v1%, _
            po1, po2, rela)
   If out% = -1 Then
    combine_two_polygon0 = -2
   ElseIf out% = -2 Then
    combine_two_polygon0 = -1
   ElseIf out% = -3 Then
    combine_two_polygon0 = -4
   ElseIf out% = -4 Then
    combine_two_polygon0 = -3
   ElseIf out% = -5 Then
    combine_two_polygon0 = -6
   ElseIf out% = -6 Then
    combine_two_polygon0 = -5
   Else
    combine_two_polygon0 = out%
   End If
  Else 'If v1% = 4 And v2% = 3 Then
    If tl(0) = tl(1) And tl(2) = tl(3) Then
     combine_two_polygon0 = combine_two_polygon0_with_3line(p1%, p2%, p3%, p4%, _
      p5%, p6%, p6%, p8%, tn1(0), tn1(1), tn1(2), tn1(3), tn2(0), tn2(1), tn2(2), tn2(3), _
          po1, po2)
         Exit Function
    ElseIf p1% = p5% Then
     combine_two_polygon0 = combine_two_polygon0_with_3point(p1%, p2%, p3%, p4%, p5%, p6%, p7%, p8%, _
          po1, po2)
           Exit Function
    ElseIf p4% = p8% Then
     combine_two_polygon0 = combine_two_polygon0_with_3point(p4%, p3%, p2%, p1%, p8%, p7%, p6%, p5%, _
          po1, po2)
           Exit Function
    End If
  End If
End Function
Public Function combine_midpoint_with_midpoint(ByVal i%, _
      ByVal no_reduce As Byte) As Byte
Dim temp_record As total_record_type
Dim re As record_data_type
Dim ty, ty1 As Byte
Dim k%, l%, tn%, no%, j%, t%
Dim n(2) As Integer
Dim m(2) As Integer
Dim dn(2) As Integer
Dim tl(2) As Integer
Dim n_(1) As Integer
Dim last_tn1%, last_tn2%
Dim t_n1() As Integer
Dim t_n2() As Integer
Dim tn_(3) As Integer
Dim con_ty As Byte
Dim ty0 As Boolean
Dim md_ As mid_point_data0_type
For k% = 0 To 2
 For l% = 0 To 2
  combine_midpoint_with_midpoint = _
      combine_mid_point_with_mid_point0(i%, k%, l%, no_reduce)
   If combine_midpoint_with_midpoint > 1 Then
    Exit Function
   End If
Next l%
Next k%
re.data0.theorem_no = 1
Call add_conditions_to_record(midpoint_, i%, 0, 0, re.data0.condition_data)
'***********
For t% = 1 + last_conditions.last_cond(0).mid_point_no To last_conditions.last_cond(1).mid_point_no
j% = Dmid_point(t%).data(0).record.data1.index.i(0)
If i% <> j% Then
  temp_record.record_data = re
   Call add_conditions_to_record(midpoint_, j%, 0, 0, temp_record.record_data.data0.condition_data) '
   temp_record.record_data.data0.theorem_no = 1
If Dmid_point(i%).data(0).data0.line_no = Dmid_point(j%).data(0).data0.line_no Then
   If Dmid_point(i%).data(0).data0.poi(2) = Dmid_point(j%).data(0).data0.poi(0) Then
      combine_midpoint_with_midpoint = set_Drelation(Dmid_point(i%).data(0).data0.poi(0), _
         Dmid_point(j%).data(0).data0.poi(2), Dmid_point(i%).data(0).data0.poi(1), _
          Dmid_point(j%).data(0).data0.poi(1), Dmid_point(i%).data(0).data0.n(0), _
           Dmid_point(j%).data(0).data0.n(2), Dmid_point(i%).data(0).data0.n(1), _
            Dmid_point(j%).data(0).data0.n(1), Dmid_point(i%).data(0).data0.line_no, _
             Dmid_point(i%).data(0).data0.line_no, "2", temp_record, 0, 0, 0, 0, 0, False)
      If combine_midpoint_with_midpoint > 1 Then
         Exit Function
      End If
   ElseIf Dmid_point(i%).data(0).data0.poi(0) = Dmid_point(j%).data(0).data0.poi(2) Then
      combine_midpoint_with_midpoint = set_Drelation(Dmid_point(j%).data(0).data0.poi(0), _
         Dmid_point(i%).data(0).data0.poi(2), Dmid_point(j%).data(0).data0.poi(1), _
          Dmid_point(i%).data(0).data0.poi(1), Dmid_point(j%).data(0).data0.n(0), _
           Dmid_point(i%).data(0).data0.n(2), Dmid_point(j%).data(0).data0.n(1), _
            Dmid_point(i%).data(0).data0.n(1), Dmid_point(i%).data(0).data0.line_no, _
             Dmid_point(i%).data(0).data0.line_no, "2", temp_record, 0, 0, 0, 0, 0, False)
      If combine_midpoint_with_midpoint > 1 Then
         Exit Function
      End If
   ElseIf Dmid_point(i%).data(0).data0.poi(0) = Dmid_point(j%).data(0).data0.poi(0) Then
      combine_midpoint_with_midpoint = set_Drelation(Dmid_point(i%).data(0).data0.poi(2), _
         Dmid_point(j%).data(0).data0.poi(2), Dmid_point(i%).data(0).data0.poi(1), _
          Dmid_point(j%).data(0).data0.poi(1), Dmid_point(i%).data(0).data0.n(2), _
           Dmid_point(j%).data(0).data0.n(2), Dmid_point(i%).data(0).data0.n(1), _
            Dmid_point(j%).data(0).data0.n(1), Dmid_point(i%).data(0).data0.line_no, _
             Dmid_point(i%).data(0).data0.line_no, "2", temp_record, 0, 0, 0, 0, 0, False)
      If combine_midpoint_with_midpoint > 1 Then
         Exit Function
      End If
   ElseIf Dmid_point(i%).data(0).data0.poi(1) = Dmid_point(j%).data(0).data0.poi(1) Then
    combine_midpoint_with_midpoint = set_equal_dline( _
        Dmid_point(i%).data(0).data0.poi(2), Dmid_point(j%).data(0).data0.poi(2), _
          Dmid_point(i%).data(0).data0.poi(0), Dmid_point(j%).data(0).data0.poi(0), _
           0, 0, 0, 0, 0, 0, 0, temp_record, 0, 0, 0, 0, 0, False)
      If combine_midpoint_with_midpoint > 1 Then
         Exit Function
      End If
    combine_midpoint_with_midpoint = set_equal_dline( _
        Dmid_point(i%).data(0).data0.poi(2), Dmid_point(j%).data(0).data0.poi(0), _
         Dmid_point(i%).data(0).data0.poi(0), Dmid_point(j%).data(0).data0.poi(2), _
          0, 0, 0, 0, 0, 0, 0, temp_record, 0, 0, 0, 0, 0, False)
      If combine_midpoint_with_midpoint > 1 Then
         Exit Function
      End If
   ElseIf Dmid_point(i%).data(0).data0.poi(2) = Dmid_point(j%).data(0).data0.poi(2) Then
      combine_midpoint_with_midpoint = set_Drelation(Dmid_point(i%).data(0).data0.poi(0), _
         Dmid_point(j%).data(0).data0.poi(0), Dmid_point(i%).data(0).data0.poi(1), _
          Dmid_point(j%).data(0).data0.poi(1), Dmid_point(i%).data(0).data0.n(0), _
           Dmid_point(j%).data(0).data0.n(0), Dmid_point(i%).data(0).data0.n(1), _
            Dmid_point(j%).data(0).data0.n(1), Dmid_point(i%).data(0).data0.line_no, _
             Dmid_point(i%).data(0).data0.line_no, "2", temp_record, 0, 0, 0, 0, 0, False)
      If combine_midpoint_with_midpoint > 1 Then
         Exit Function
      End If
   End If
Else
  temp_record.record_data = re
   Call add_conditions_to_record(midpoint_, j%, 0, 0, temp_record.record_data.data0.condition_data) '
If is_dparal(line_number0(Dmid_point(i%).data(0).data0.poi(0), Dmid_point(j%).data(0).data0.poi(0), _
    tn_(0), tn_(1)), line_number0(Dmid_point(i%).data(0).data0.poi(2), Dmid_point(j%).data(0).data0.poi(2), _
      tn_(2), tn_(3)), tn%, -1000, 0, 0, 0, 0) And _
       th_chose(98).chose = 1 Then
 'If tn% > 0 Then
     temp_record.record_data.data0.theorem_no = 98
Call add_conditions_to_record(paral_, tn%, 0, 0, temp_record.record_data.data0.condition_data)
combine_midpoint_with_midpoint = _
 set_dparal(line_number0(Dmid_point(i%).data(0).data0.poi(2), _
            Dmid_point(j%).data(0).data0.poi(2), 0, 0), _
              line_number0(Dmid_point(i%).data(0).data0.poi(1), _
                Dmid_point(j%).data(0).data0.poi(1), 0, 0), _
                 temp_record, 0, no_reduce, False)
  If combine_midpoint_with_midpoint > 1 Then
   Exit Function
  End If
combine_midpoint_with_midpoint = _
 set_dparal(line_number0(Dmid_point(i%).data(0).data0.poi(0), _
             Dmid_point(j%).data(0).data0.poi(0), 0, 0), _
              line_number0(Dmid_point(i%).data(0).data0.poi(1), _
                Dmid_point(j%).data(0).data0.poi(1), 0, 0), _
                  temp_record, 0, no_reduce, False)
  If combine_midpoint_with_midpoint > 1 Then
   Exit Function
  End If
  If (tn_(0) > tn_(1) And tn_(2) > tn_(3)) Or (tn_(0) < tn_(1) And tn_(2) < tn_(3)) Then
   combine_midpoint_with_midpoint = _
    set_three_line_value(Dmid_point(i%).data(0).data0.poi(0), Dmid_point(j%).data(0).data0.poi(0), _
     Dmid_point(j%).data(0).data0.poi(2), Dmid_point(i%).data(0).data0.poi(2), Dmid_point(i%).data(0).data0.poi(1), _
      Dmid_point(j%).data(0).data0.poi(1), 0, 0, 0, 0, 0, 0, 0, 0, 0, _
       "1", "1", "-2", "0", temp_record, 0, no_reduce, 0)
  If combine_midpoint_with_midpoint > 1 Then
   Exit Function
  End If
  Else
    If squre_distance_point_point(m_poi(Dmid_point(i%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                  m_poi(Dmid_point(j%).data(0).data0.poi(0)).data(0).data0.coordinate) > _
        squre_distance_point_point(m_poi(Dmid_point(i%).data(0).data0.poi(2)).data(0).data0.coordinate, _
                       m_poi(Dmid_point(j%).data(0).data0.poi(2)).data(0).data0.coordinate) Then
   combine_midpoint_with_midpoint = _
    set_three_line_value(Dmid_point(i%).data(0).data0.poi(0), Dmid_point(j%).data(0).data0.poi(0), _
     Dmid_point(j%).data(0).data0.poi(2), Dmid_point(i%).data(0).data0.poi(2), Dmid_point(i%).data(0).data0.poi(1), _
      Dmid_point(j%).data(0).data0.poi(1), 0, 0, 0, 0, 0, 0, 0, 0, 0, _
       "1", "-1", "-2", "0", temp_record, 0, no_reduce, 0)
    If combine_midpoint_with_midpoint > 1 Then
     Exit Function
    End If
   Else
   combine_midpoint_with_midpoint = _
    set_three_line_value(Dmid_point(i%).data(0).data0.poi(0), Dmid_point(j%).data(0).data0.poi(0), _
     Dmid_point(j%).data(0).data0.poi(2), Dmid_point(i%).data(0).data0.poi(2), Dmid_point(i%).data(0).data0.poi(1), _
      Dmid_point(j%).data(0).data0.poi(1), 0, 0, 0, 0, 0, 0, 0, 0, 0, _
       "-1", "1", "-2", "0", temp_record, 0, no_reduce, 0)
    If combine_midpoint_with_midpoint > 1 Then
     Exit Function
    End If
   End If
  End If
 ElseIf is_dparal(line_number0(Dmid_point(i%).data(0).data0.poi(0), Dmid_point(j%).data(0).data0.poi(2), tn_(0), tn_(1)), _
                   line_number0(Dmid_point(i%).data(0).data0.poi(2), Dmid_point(j%).data(0).data0.poi(0), tn_(2), tn_(3)), _
                    tn%, -1000, 0, 0, 0, 0) And th_chose(89).chose = 1 Then
   If tn% > 0 Then
    Call add_conditions_to_record(paral_, tn%, 0, 0, temp_record.record_data.data0.condition_data)
   End If
   temp_record.record_data.data0.theorem_no = 89
  combine_midpoint_with_midpoint = set_dparal(line_number0(Dmid_point(i%).data(0).data0.poi(1), _
       Dmid_point(j%).data(0).data0.poi(1), 0, 0), line_number0(Dmid_point(i%).data(0).data0.poi(0), _
         Dmid_point(j%).data(0).data0.poi(2), 0, 0), temp_record, 0, no_reduce, False)
  If combine_midpoint_with_midpoint > 1 Then
   Exit Function
  End If
  combine_midpoint_with_midpoint = set_dparal(line_number0(Dmid_point(i%).data(0).data0.poi(1), _
       Dmid_point(j%).data(0).data0.poi(1), 0, 0), line_number0(Dmid_point(i%).data(0).data0.poi(2), _
         Dmid_point(j%).data(0).data0.poi(0), 0, 0), temp_record, 0, no_reduce, False)
  If combine_midpoint_with_midpoint > 1 Then
   Exit Function
  End If
   If (tn_(0) > tn_(1) And tn_(2) > tn_(3)) Or (tn_(0) < tn_(1) And tn_(2) < tn_(3)) Then
   combine_midpoint_with_midpoint = _
    set_three_line_value(Dmid_point(i%).data(0).data0.poi(0), Dmid_point(j%).data(0).data0.poi(2), _
     Dmid_point(j%).data(0).data0.poi(2), Dmid_point(i%).data(0).data0.poi(0), Dmid_point(i%).data(0).data0.poi(1), _
      Dmid_point(j%).data(0).data0.poi(1), 0, 0, 0, 0, 0, 0, 0, 0, 0, _
       "1", "1", "-2", "0", temp_record, 0, no_reduce, 0)
  If combine_midpoint_with_midpoint > 1 Then
   Exit Function
  End If
  Else
    If squre_distance_point_point(m_poi(Dmid_point(i%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                  m_poi(Dmid_point(j%).data(0).data0.poi(2)).data(0).data0.coordinate) > _
        squre_distance_point_point(m_poi(Dmid_point(i%).data(0).data0.poi(2)).data(0).data0.coordinate, _
                       m_poi(Dmid_point(j%).data(0).data0.poi(0)).data(0).data0.coordinate) Then
   combine_midpoint_with_midpoint = _
    set_three_line_value(Dmid_point(i%).data(0).data0.poi(0), Dmid_point(j%).data(0).data0.poi(2), _
     Dmid_point(i%).data(0).data0.poi(2), Dmid_point(j%).data(0).data0.poi(0), Dmid_point(i%).data(0).data0.poi(1), _
      Dmid_point(j%).data(0).data0.poi(1), 0, 0, 0, 0, 0, 0, 0, 0, 0, _
       "1", "-1", "-2", "0", temp_record, 0, no_reduce, 0)
  If combine_midpoint_with_midpoint > 1 Then
   Exit Function
  End If
  Else
   combine_midpoint_with_midpoint = _
    set_three_line_value(Dmid_point(i%).data(0).data0.poi(0), Dmid_point(j%).data(0).data0.poi(2), _
     Dmid_point(i%).data(0).data0.poi(2), Dmid_point(j%).data(0).data0.poi(0), Dmid_point(i%).data(0).data0.poi(1), _
      Dmid_point(j%).data(0).data0.poi(1), 0, 0, 0, 0, 0, 0, 0, 0, 0, _
       "-1", "1", "-2", "0", temp_record, 0, no_reduce, 0)
  If combine_midpoint_with_midpoint > 1 Then
   Exit Function
  End If
  End If
 End If
 End If
End If
End If
Next t%
'**********************************************************************

   '对角线相互平分的四边形是平行四边形
  md_.poi(m(0)) = Dmid_point(i%).data(0).data0.poi(1)
  md_.poi(m(1)) = -1
  Call search_for_mid_point(md_, 1, n_(0), 1)
  md_.poi(m(1)) = 30000
  Call search_for_mid_point(md_, 1, n_(1), 1)  '5.7
  For j% = n_(0) + 1 To n_(1)
  no% = Dmid_point(j%).data(0).record.data1.index.i(1)
  If Dmid_point(i%).data(0).data0.line_no <> Dmid_point(no%).data(0).data0.line_no Then
   If no% > 0 And no% < i% Then
      last_tn1% = last_tn1% + 1
       ReDim Preserve t_n1(last_tn1%) As Integer
        t_n1(last_tn1%) = no%
     End If
   End If
   Next j%
  For j% = 1 To last_tn1%
  no% = t_n1(j%)
   temp_record.record_data = re
    Call add_conditions_to_record(midpoint_, no%, 0, 0, temp_record.record_data.data0.condition_data)
  temp_record.record_data.data0.theorem_no = 64
  combine_midpoint_with_midpoint = _
   set_parallelogram(Dmid_point(i%).data(0).data0.poi(0), _
    Dmid_point(no%).data(0).data0.poi(0), Dmid_point(i%).data(0).data0.poi(2), _
     Dmid_point(no%).data(0).data0.poi(2), temp_record, 0, no_reduce)
  If combine_midpoint_with_midpoint > 1 Then
   Exit Function
  End If
 Next j%
End Function
Public Function combine_relation_with_item(ByVal re%, no_reduce As Byte) As Byte
Dim i%, j%, k%, no%
Dim n_(1) As Integer
Dim n(2) As Integer
Dim m(2) As Integer
Dim m1(2) As Integer
Dim it As item0_data_type
Dim tn() As Integer
Dim last_tn%
If Drelation(re%).record_.no_reduce > 4 Then
 Exit Function
End If
For i% = 0 To 2
 n(0) = i%
  n(1) = (i% + 1) Mod 3
   n(2) = (i% + 2) Mod 3
For j% = 0 To 2
 m(0) = j%
  m(1) = (j% + 1) Mod 3
   m(2) = (j% + 2) Mod 3
it.poi(2 * m(0)) = Drelation(re%).data(0).data0.poi(2 * n(0))
it.poi(2 * m(0) + 1) = Drelation(re%).data(0).data0.poi(2 * n(0) + 1)
it.poi(2 * m(1)) = -1
Call search_for_item0(it, m(0), n_(0), 1)   '5.7
it.poi(2 * m(1)) = 30000
Call search_for_item0(it, m(0), n_(1), 1)
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
 no% = item0(k%).data(0).index(m(0))
 If no% > 0 Then
    last_tn% = last_tn% + 1
     ReDim Preserve tn(last_tn%) As Integer
    tn(last_tn%) = no%
 End If
Next k%
 For k% = 1 To last_tn%
 no% = tn(k%)
  combine_relation_with_item = combine_relation_with_item_(relation_, re%, _
    no%, i%, j%, no_reduce)
   If combine_relation_with_item > 1 Then
    Exit Function
   End If
 Next k%
Next j%
Next i%
End Function
Public Function combine_three_angle_with_three_angle_( _
           ByVal A3%, no%, no_reduce As Byte) As Byte
           'combine_angle
Dim i%, j%, k%, l%, tn%, no_%
Dim n1_(2) As Integer
Dim n2_(2) As Integer
Dim tA(3) As Integer
Dim tA1(2) As Integer
Dim tA2(2) As Integer
Dim s1(2) As String
Dim v(1) As String
Dim S2(2) As String
Dim ty As Byte
Dim t_A As angle3_value_data0_type
Dim t_A1 As angle3_value_data0_type
Dim temp_record As total_record_type
     temp_record.record_data.data0.condition_data.condition_no = 2
      temp_record.record_data.data0.condition_data.condition(1).ty = angle3_value_
       temp_record.record_data.data0.condition_data.condition(1).no = A3%
      temp_record.record_data.data0.condition_data.condition(2).ty = angle3_value_
       temp_record.record_data.data0.condition_data.condition(2).no = no%
        temp_record.record_data.data0.theorem_no = 1
   ' 共边角合并
   For j% = 0 To 2
    For k% = 0 To 2
     If angle3_value(A3%).data(0).data0.angle(j%) > 0 And _
          angle3_value(no%).data(0).data0.angle(k%) > 0 Then
      If angle(angle3_value(A3%).data(0).data0.angle(j%)).data(0).line_no(0) = _
          angle(angle3_value(no%).data(0).data0.angle(k%)).data(0).line_no(0) Or _
        angle(angle3_value(A3%).data(0).data0.angle(j%)).data(0).line_no(0) = _
          angle(angle3_value(no%).data(0).data0.angle(k%)).data(0).line_no(1) Or _
        angle(angle3_value(A3%).data(0).data0.angle(j%)).data(0).line_no(1) = _
          angle(angle3_value(no%).data(0).data0.angle(k%)).data(0).line_no(0) Or _
        angle(angle3_value(A3%).data(0).data0.angle(j%)).data(0).line_no(1) = _
          angle(angle3_value(no%).data(0).data0.angle(k%)).data(0).line_no(1) Then
         n1_(0) = j%
          n2_(0) = k%
         n1_(1) = (j% + 1) Mod 3
          n2_(1) = (k% + 1) Mod 3
         n1_(2) = (j% + 2) Mod 3
          n2_(2) = (k% + 2) Mod 3
           If combine_two_angle(angle3_value(A3%).data(0).data0.angle(j%), _
                angle3_value(no%).data(0).data0.angle(k%), tA(0), 0, tA(3), tA(1), 0, tA(2), ty, 0, 1) Then
             '*****************
             tA1(0) = angle3_value(A3%).data(0).data0.angle(n1_(0))
              tA1(1) = angle3_value(A3%).data(0).data0.angle(n1_(1))
               tA1(2) = angle3_value(A3%).data(0).data0.angle(n1_(2))
             s1(0) = time_string(angle3_value(A3%).data(0).data0.para(n1_(0)), _
                    angle3_value(no%).data(0).data0.para(n2_(0)), True, False)
              s1(1) = time_string(angle3_value(A3%).data(0).data0.para(n1_(1)), _
                    angle3_value(no%).data(0).data0.para(n2_(0)), True, False)
               s1(2) = time_string(angle3_value(A3%).data(0).data0.para(n1_(2)), _
                  angle3_value(no%).data(0).data0.para(n2_(0)), True, False)
                v(0) = time_string(angle3_value(A3%).data(0).data0.value, _
                  angle3_value(no%).data(0).data0.para(n2_(0)), True, False)
             tA2(0) = angle3_value(no%).data(0).data0.angle(n2_(0))
              tA2(1) = angle3_value(no%).data(0).data0.angle(n2_(1))
               tA2(2) = angle3_value(no%).data(0).data0.angle(n2_(2))
             S2(0) = time_string(angle3_value(no%).data(0).data0.para(n2_(0)), _
                      angle3_value(A3%).data(0).data0.para(n1_(0)), True, False)
              S2(1) = time_string(angle3_value(no%).data(0).data0.para(n2_(1)), _
                angle3_value(A3%).data(0).data0.para(n1_(0)), True, False)
               S2(2) = time_string(angle3_value(no%).data(0).data0.para(n2_(2)), _
                  angle3_value(A3%).data(0).data0.para(n1_(0)), True, False)
                v(1) = time_string(angle3_value(no%).data(0).data0.value, _
                     angle3_value(A3%).data(0).data0.para(n1_(0)), True, False)
              '******************
            If ty = 2 Then
              tA1(0) = 0
               s1(0) = "0"
              tA2(0) = 0
               S2(0) = "0"
                S2(1) = time_string(S2(1), "-1", True, False)
                 S2(2) = time_string(S2(2), "-1", True, False)
                  v(1) = time_string(v(1), "-1", True, False)
            ElseIf ty = 3 Or ty = 5 Then
              tA2(0) = 0
               S2(0) = "0"
                tA1(0) = tA(2)
            ElseIf ty = 6 Or ty = 4 Then
              If ty = 6 Then
               s1(0) = time_string(s1(0), "-1", True, False)
              End If
              tA1(0) = tA(0)
              tA2(0) = 0
               S2(0) = "0"
                S2(1) = time_string(S2(1), "-1", True, False)
                 S2(2) = time_string(S2(2), "-1", True, False)
                  v(1) = time_string(v(1), "-1", True, False)
            ElseIf ty = 7 Or ty = 8 Then
             tA1(0) = tA(1)
              If ty = 7 Then
               s1(0) = time_string(s1(0), "-1", True, False)
              End If
              tA2(0) = 0
               S2(0) = "0"
                S2(1) = time_string(S2(1), "-1", True, False)
                 S2(2) = time_string(S2(2), "-1", True, False)
                  v(1) = time_string(v(1), "-1", True, False)
            ElseIf ty = 10 Or ty = 9 Then
             tA2(0) = 0
             S2(0) = "0"
             tA1(0) = tA(2)
             s1(0) = time_string(s1(0), "-1", True, False)
             v(1) = add_string(v(1), time_string(s1(0), "360", False, False), True, False)
            ElseIf ty = 11 Or ty = 15 Then
             tA1(0) = tA(3)
              v(0) = minus_string(v(0), time_string("180", s1(0), False, False), True, False)
               s1(0) = time_string(s1(0), "-1", True, False)
               tA2(0) = 0
                S2(0) = "0"
            ElseIf ty = 12 Or ty = 16 Then
             tA1(0) = tA(3)
               tA2(0) = 0
                S2(0) = "0"
              v(0) = minus_string(v(0), time_string("180", s1(0), False, False), True, False)
            ElseIf ty = 19 Or ty = 20 Then
             tA1(0) = 0
              v(0) = minus_string(v(0), time_string("180", s1(0), False, False), True, False)
               s1(0) = "0"
               tA2(0) = 0
                S2(0) = "0"
            End If
             combine_three_angle_with_three_angle_ = combine_six_angle0( _
               tA1(0), tA1(1), tA1(2), tA2(0), tA2(1), tA2(2), _
                s1(0), s1(1), s1(2), S2(0), S2(1), S2(2), _
                 add_string(v(0), v(1), True, False), 5, temp_record.record_data)
         End If
        End If
     End If
    Next k%
 Next j%
End Function
Public Function combine_three_angle_with_three_angle0( _
   A3 As angle3_value_data0_type, re As record_data_type) As Byte
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
temp_record.record_data = re
temp_record.record_data.data0.theorem_no = 1
If temp_record.record_data.data0.condition_data.condition(2).no > 0 Then
If angle3_value(temp_record.record_data.data0.condition_data.condition(2).no).data(0).data0.para(2) <> "0" Then
tA% = temp_record.record_data.data0.condition_data.condition(2).no
Else
tA% = temp_record.record_data.data0.condition_data.condition(1).no
End If
Else
 tA% = temp_record.record_data.data0.condition_data.condition(1).no
End If
For i% = 0 To 3
If A3.angle(i%) > 0 Then
'If angle3_value(A3%).data.para(2) <> "0" Then
If i% < 3 Then
 n(0) = i%
  n(1) = (i% + 1) Mod 3
   n(2) = (i% + 2) Mod 3
Else
 n(0) = 3
  n(1) = 0
End If
For j% = 0 To 3
If j% < 3 Then
 m(0) = j%
  m(1) = (j% + 1) Mod 3
   m(2) = (j% + 2) Mod 3
Else
 m(0) = 3
  m(1) = 0
End If
t_A.angle(m(0)) = A3.angle(n(0))
t_A.angle(m(1)) = -1
Call search_for_three_angle_value(t_A, j%, n_(0), 1)   '5.7
t_A.angle(m(1)) = 30000
Call search_for_three_angle_value(t_A, j%, n_(1), 1)   '5.7
last_tn% = 0
last_tn1% = 0
last_tn2% = 0
last_tn3% = 0
For k% = n_(0) + 1 To n_(1)
no% = angle3_value(k%).data(0).record.data1.index.i(j%)
If m(0) < 3 And n(0) < 3 Then
If angle3_value(no%).data(0).data0.type = angle_value_ Then
 GoTo combine_three_angle_with_three_angle0_mark0
ElseIf (A3.type = angle_value_ And angle3_value(no%).data(0).data0.type = angle_value_) Or _
        (A3.type = eangle_ And angle3_value(no%).data(0).data0.type = eangle_) Then
    GoTo combine_three_angle_with_three_angle0_mark0
   End If
End If
If no% > 0 And no% <> re.data0.condition_data.condition(1).no And no% < tA% And angle3_value(no%).data(0).data0.reduce Then
' If angle3_value(no%).data(0).data0.para(1) <> "0" Then
   'If is_two_record_related(angle3_value_, no%, angle3_value(no%).data(0).record, _
       angle3_value_, re.data0.condition_data.condition(1).no, angle3_value(re.data0.condition_data.condition(1).no).data(0).record) = False Then
  If angle3_value(no%).record_.no_reduce < 255 Then
If no% < re.data0.condition_data.condition(1).no Then
  If re.data0.condition_data.condition_no = 2 Then
  If no% > re.data0.condition_data.condition(2).no Or angle3_value(no%).data(0).data0.value <> "0" Then
       GoTo combine_three_angle_with_three_angle0_mark0
  Else
  'If is_two_record_related(re.data0.condition_data.condition(2).ty, re.data0.condition_data.condition(2).no, _
       angle3_value(re.data0.condition_data.condition(2).no).data(0).record, angle3_value_, no%, _
         angle3_value(no%).data(0).record) Then
      ' GoTo combine_three_angle_with_three_angle0_mark0
 ' End If
  If no% > re.data0.condition_data.condition(2).no Or angle3_value(no%).data(0).data0.para(2) <> "0" Or _
       angle3_value(no%).data(0).data0.value <> "0" Then
       GoTo combine_three_angle_with_three_angle0_mark0
  End If
  End If
' End If
 End If
If n(0) < 3 And m(0) < 3 Then '主数据
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
ElseIf m(0) < 3 Then 'n(0)=3
last_tn1% = last_tn1% + 1
ReDim Preserve tn1(last_tn1%) As Integer
tn1(last_tn1%) = no%
ElseIf n(0) < 3 Then 'm(0)=3
last_tn2% = last_tn2% + 1
ReDim Preserve tn2(last_tn2%) As Integer
tn2(last_tn2%) = no%
Else 'n(0)=3 and m(0)=3
last_tn3% = last_tn3% + 1
ReDim Preserve tn3(last_tn3%) As Integer
tn3(last_tn3%) = no%
End If
End If
combine_three_angle_with_three_angle_mark1:
End If
End If
'End If
combine_three_angle_with_three_angle0_mark0:
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
  temp_record.record_data = re
  'temp_record.data.data0.theorem_no = 1
  Call add_conditions_to_record(angle3_value_, no%, 0, 0, temp_record.record_data.data0.condition_data)
  tA1(0) = A3.angle(n(1))
  tA1(1) = A3.angle(n(2))
  tA2(0) = angle3_value(no%).data(0).data0.angle(m(1))
  tA2(1) = angle3_value(no%).data(0).data0.angle(m(2))
  s1(0) = A3.para(n(0))
  s1(1) = A3.para(n(1))
  s1(2) = A3.para(n(2))
  S2(0) = angle3_value(no%).data(0).data0.para(m(0))
  S2(1) = angle3_value(no%).data(0).data0.para(m(1))
  S2(2) = angle3_value(no%).data(0).data0.para(m(2))
  v(0) = A3.value
  v(1) = angle3_value(no%).data(0).data0.value
  combine_three_angle_with_three_angle0 = _
   solve_equation_for_angle3(tA%, no%, A3.angle(n(0)), _
     tA1(0), tA1(1), s1(0), s1(1), s1(2), v(0), _
      tA2(0), tA2(1), S2(0), S2(1), S2(2), v(1), temp_record.record_data.data0.condition_data)
   'If combine_three_angle_with_three_angle0 > 0 Then
    'angle3_value(no%).data(0).data0.reduce = False
   'End If
   If combine_three_angle_with_three_angle0 > 1 Then
    Exit Function
   End If
Next k%
For k% = 1 To last_tn1%
 no% = tn1(k%)
  temp_record.record_data = re
  Call add_conditions_to_record(angle3_value_, no%, 0, 0, temp_record.record_data.data0.condition_data)
  tA2(0) = angle3_value(no%).data(0).data0.angle(m(1))
  tA2(1) = angle3_value(no%).data(0).data0.angle(m(2))
  S2(0) = angle3_value(no%).data(0).data0.para(m(0))
  S2(1) = angle3_value(no%).data(0).data0.para(m(1))
  S2(2) = angle3_value(no%).data(0).data0.para(m(2))
  v(1) = angle3_value(no%).data(0).data0.value
  Call read_angle_para_from_angle3_value(A3, 3, 0, tA1(0), tA1(1), s1(0), s1(1), s1(2), v(0), 0)
 combine_three_angle_with_three_angle0 = _
   solve_equation_for_angle3(tA%, no%, A3.angle(n(0)), _
     tA1(0), tA1(1), s1(0), s1(1), s1(2), v(0), _
      tA2(0), tA2(1), S2(0), S2(1), S2(2), v(1), temp_record.record_data.data0.condition_data)
   If combine_three_angle_with_three_angle0 > 1 Then
    Exit Function
   End If
   Call read_angle_para_from_angle3_value(A3, 3, 0, tA1(0), tA1(1), s1(0), s1(1), s1(2), v(0), 1)
combine_three_angle_with_three_angle0 = _
   solve_equation_for_angle3(tA%, no%, A3.angle(n(0)), _
      tA1(0), tA1(1), s1(0), s1(1), s1(2), v(0), _
       tA2(0), tA2(1), S2(0), S2(1), S2(2), v(1), temp_record.record_data.data0.condition_data)
   'If combine_three_angle_with_three_angle0 > 0 Then
   ' angle3_value(no%).data(0).data0.reduce = False
   'End If
   If combine_three_angle_with_three_angle0 > 1 Then
    Exit Function
   End If
Next k%
For k% = 1 To last_tn2%
 no% = tn2(k%)
  temp_record.record_data = re
  'temp_record.data.data0.theorem_no = 1
  Call add_conditions_to_record(angle3_value_, no%, 0, 0, temp_record.record_data.data0.condition_data)
  tA1(0) = A3.angle(n(1))
  tA1(1) = A3.angle(n(2))
  'tA2(0) = angle3_value(no%).data(0).data0.angle(1)
  'tA2(1) = angle3_value(no%).data(0).data0.angle(2)
  s1(0) = A3.para(n(0))
  s1(1) = A3.para(n(1))
  s1(2) = A3.para(n(2))
  'S2(0) = angle3_value(no%).data(0).data0.para(0)
  'S2(1) = minus_string(angle3_value(no%).data(0).data0.para(1), _
            angle3_value(no%).data(0).data0.para(0), True, False)
   'If S2(1) = "0" Then
    'tA2(1) = 0
   'End If
  'S2(2) = angle3_value(no%).data(0).data0.para(2)
  v(0) = A3.value
  'v(1) = angle3_value(no%).data(0).data0.value
    Call read_angle_para_from_angle3_value(angle3_value(no%).data(0).data0, 3, 0, tA2(0), tA2(1), S2(0), _
         S2(1), S2(2), v(1), 0)
  combine_three_angle_with_three_angle0 = _
   solve_equation_for_angle3(tA%, no%, A3.angle(n(0)), _
      tA1(0), tA1(1), s1(0), s1(1), s1(2), v(0), _
       tA2(0), tA2(1), S2(0), S2(1), S2(2), v(1), temp_record.record_data.data0.condition_data)
   'If combine_three_angle_with_three_angle0 > 0 Then
   ' angle3_value(no%).data(0).data0.reduce = False
   'End If
   If combine_three_angle_with_three_angle0 > 1 Then
    Exit Function
   End If
  'tA1(0) = A3.angle(n(1))
  'tA1(1) = A3.angle(n(2))
  'tA2(0) = angle3_value(no%).data(0).data0.angle(0)
  'tA2(1) = angle3_value(no%).data(0).data0.angle(2)
  's1(0) = A3.para(n(0))
  's1(1) = A3.para(n(1))
  's1(2) = A3.para(n(2))
  'S2(0) = angle3_value(no%).data(0).data0.para(1)
  'S2(1) = minus_string(angle3_value(no%).data(0).data0.para(0), _
            angle3_value(no%).data(0).data0.para(1), True, False)
   'If S2(1) = "0" Then
    'tA2(1) = 0
   'End If
  'S2(2) = angle3_value(no%).data(0).data0.para(2)
  'v(0) = A3.value
  'v(1) = angle3_value(no%).data(0).data0.value
     Call read_angle_para_from_angle3_value(angle3_value(no%).data(0).data0, 3, 0, tA2(0), tA2(1), S2(0), _
         S2(1), S2(2), v(1), 1)
 combine_three_angle_with_three_angle0 = _
   solve_equation_for_angle3(tA%, no%, A3.angle(n(0)), _
      tA1(0), tA1(1), s1(0), s1(1), s1(2), v(0), _
       tA2(0), tA2(1), S2(0), S2(1), S2(2), v(1), temp_record.record_data.data0.condition_data)
    If combine_three_angle_with_three_angle0 > 1 Then
    Exit Function
   End If
Next k%
For k% = 1 To last_tn3%
 no% = tn3(k%)
  temp_record.record_data = re
  'temp_record.data.data0.theorem_no = 1
  Call add_conditions_to_record(angle3_value_, no%, 0, 0, temp_record.record_data.data0.condition_data)
    Call read_angle_para_from_angle3_value(angle3_value(tA%).data(0).data0, 3, 0, tA1(0), tA1(1), s1(0), _
         s1(1), s1(2), v(0), 0)
     Call read_angle_para_from_angle3_value(angle3_value(no%).data(0).data0, 3, 0, tA2(0), tA2(1), S2(0), _
         S2(1), S2(2), v(1), 0)
 combine_three_angle_with_three_angle0 = _
   solve_equation_for_angle3(tA%, no%, A3.angle(n(0)), _
      tA1(0), tA1(1), s1(0), s1(1), s1(2), v(0), _
       tA2(0), tA2(1), S2(0), S2(1), S2(2), v(1), temp_record.record_data.data0.condition_data)
    If combine_three_angle_with_three_angle0 > 1 Then
    Exit Function
   End If
     Call read_angle_para_from_angle3_value(angle3_value(no%).data(0).data0, 3, 0, tA2(0), tA2(1), S2(0), _
         S2(1), S2(2), v(1), 1)
 combine_three_angle_with_three_angle0 = _
   solve_equation_for_angle3(tA%, no%, A3.angle(n(0)), _
      tA1(0), tA1(1), s1(0), s1(1), s1(2), v(0), _
       tA2(0), tA2(1), S2(0), S2(1), S2(2), v(1), temp_record.record_data.data0.condition_data)
    If combine_three_angle_with_three_angle0 > 1 Then
    Exit Function
   End If
    Call read_angle_para_from_angle3_value(angle3_value(tA%).data(0).data0, 3, 0, tA1(0), tA1(1), s1(0), _
         s1(1), s1(2), v(0), 1)
 combine_three_angle_with_three_angle0 = _
   solve_equation_for_angle3(tA%, no%, A3.angle(n(0)), _
      tA1(0), tA1(1), s1(0), s1(1), s1(2), v(0), _
       tA2(0), tA2(1), S2(0), S2(1), S2(2), v(1), temp_record.record_data.data0.condition_data)
    If combine_three_angle_with_three_angle0 > 1 Then
    Exit Function
   End If
      Call read_angle_para_from_angle3_value(angle3_value(no%).data(0).data0, 3, 0, tA2(0), tA2(1), S2(0), _
         S2(1), S2(2), v(1), 0)
combine_three_angle_with_three_angle0 = _
   solve_equation_for_angle3(tA%, no%, A3.angle(n(0)), _
      tA1(0), tA1(1), s1(0), s1(1), s1(2), v(0), _
       tA2(0), tA2(1), S2(0), S2(1), S2(2), v(1), temp_record.record_data.data0.condition_data)
    If combine_three_angle_with_three_angle0 > 1 Then
    Exit Function
   End If
 Next k%
combine_three_angle_with_three_angle0_Mark10:
Next j%
End If
Next i%
End Function

Public Function combine_two_angle_with_two_angle(ByVal A11%, ByVal A12%, _
             ByVal S11 As String, ByVal S12 As String, ByVal v1 As String, _
               ByVal A21%, ByVal A22%, ByVal S21 As String, ByVal S22$, _
                ByVal v2 As String, re As record_data_type, no_reduce As Byte) As Byte
Dim ty As Byte
Dim A(2) As Integer
Dim ts(1) As String
Dim temp_record As total_record_type
temp_record.record_data = re
 ts(0) = S11
 ts(1) = S21
 Call combine_two_angle(A11%, A21%, A(0), 0, 0, A(1), 0, A(2), ty, 0, 1)
If ty = 21 Then
 ts(0) = add_string(S11, S21, True, False)
 If ts(0) = "0" Then
 combine_two_angle_with_two_angle = set_three_angle_value( _
     0, A12%, A22%, "0", S12, S22, add_string(v1, v2, True, False), _
      0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
 Else
 combine_two_angle_with_two_angle = set_three_angle_value( _
     A(0), A12%, A22%, ts(0), S12, S22, add_string(v1, v2, True, False), _
      0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
 End If
 If combine_two_angle_with_two_angle > 1 Then
  Exit Function
 End If

ElseIf ty = 3 Or ty = 5 Then
 S11 = time_string(S11, ts(1), True, False)
 S12 = time_string(S12, ts(1), True, False)
 v1 = time_string(v1, ts(1), True, False)
 S21 = time_string(S21, ts(0), True, False)
 S22 = time_string(S22, ts(0), True, False)
 v2 = time_string(v2, ts(0), True, False)
 combine_two_angle_with_two_angle = set_three_angle_value( _
     A(2), A12%, A22%, S11, S12, S22, add_string(v1, v2, True, 0), _
      0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
 If combine_two_angle_with_two_angle > 1 Then
  Exit Function
 End If
ElseIf ty = 4 Or ty = 8 Then
 S11 = time_string(S11, ts(1), True, False)
 S12 = time_string(S12, ts(1), True, False)
 v1 = time_string(v1, ts(1), True, False)
 S21 = time_string(S21, ts(0), True, False)
 S22 = time_string(S22, ts(0), True, False)
 v2 = time_string(v2, ts(0), True, False)
 If ty = 4 Then
 combine_two_angle_with_two_angle = set_three_angle_value( _
     A(0), A12%, A22%, S11, S12, time_string("-1", S22, True, False), _
      minus_string(v1, v2, True, False), 0, temp_record, 0, 0, 0, no_reduce, _
      0, 0, False)
 Else
 combine_two_angle_with_two_angle = set_three_angle_value( _
     A(1), A12%, A22%, S11, S12, time_string("-1", S22, True, False), _
      minus_string(v1, v2, True, False), 0, temp_record, 0, 0, 0, no_reduce, _
      0, 0, False)
 End If
 If combine_two_angle_with_two_angle > 1 Then
  Exit Function
 End If
ElseIf ty = 6 Or ty = 7 Then
 S11 = time_string(S11, ts(1), True, False)
 S12 = time_string(S12, ts(1), True, False)
 v1 = time_string(v1, ts(1), True, False)
 S21 = time_string(S21, ts(0), True, False)
 S22 = time_string(S22, ts(0), True, False)
 v2 = time_string(v2, ts(0), True, False)
 If ty = 6 Then
 combine_two_angle_with_two_angle = set_three_angle_value( _
     A(0), A12%, A22%, S11, time_string("-1", S12, True, False), S22, _
      minus_string(v2, v1, True, False), 0, temp_record, 0, 0, 0, no_reduce, _
      0, 0, False)
Else
 combine_two_angle_with_two_angle = set_three_angle_value( _
     A(1), A12%, A22%, S11, time_string("-1", S12, True, False), S22, _
      minus_string(v2, v1, True, False), 0, temp_record, 0, 0, 0, no_reduce, _
      0, 0, False)
End If
If combine_two_angle_with_two_angle > 1 Then
  Exit Function
 End If
 ElseIf ty = 9 Or ty = 10 Then
 S11 = time_string(S11, ts(1), True, False)
 S12 = time_string(S12, ts(1), True, False)
 v1 = time_string(v1, ts(1), True, False)
 S21 = time_string(S21, ts(0), True, False)
 S22 = time_string(S22, ts(0), True, False)
 v2 = time_string(v2, ts(0), True, False)
 combine_two_angle_with_two_angle = set_three_angle_value( _
     A(2), A12%, A22%, time_string("-1", S11, True, False), S12, S22, _
      minus_string(add_string(v1, v2, False, False), _
       time_string("360", S11, False, False), True, False), _
        0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
 If combine_two_angle_with_two_angle > 1 Then
  Exit Function
 End If
End If
End Function
Public Function combine_three_angle_with_three_angle1(tA1 As angle3_value_data0_type, _
   tA2 As angle3_value_data0_type, n11%, n12%, n21%, n22%, re As record_data_type, no_reduce) As Byte
Dim tn(1) As Integer
Dim t_n(1) As Integer
Dim ty(1) As Byte
Dim A(5) As Integer
Dim tA(1) As Integer
Dim ts(3) As String
Dim t_s$
Dim tA3_v(1) As angle3_value_data0_type
Dim temp_record As total_record_type
tA3_v(0) = tA1
tA3_v(1) = tA2
temp_record.record_data = re
If n11% = 0 Then
 If n12% = 1 Then
  t_n(0) = 2
 Else
  t_n(0) = 1
 End If
ElseIf n11% = 1 Then
 If n12% = 0 Then
  t_n(0) = 2
 Else
  t_n(0) = 0
 End If
Else
 If n12% = 0 Then
  t_n(0) = 1
 Else
  t_n(0) = 0
 End If
End If
If n21% = 0 Then
 If n22% = 1 Then
  t_n(1) = 2
 Else
  t_n(1) = 1
 End If
ElseIf n21% = 1 Then
 If n22% = 0 Then
  t_n(1) = 2
 Else
  t_n(1) = 0
 End If
Else
 If n22% = 0 Then
  t_n(1) = 1
 Else
  t_n(1) = 0
 End If
End If
 Call combine_two_angle(tA3_v(0).angle(n11%), tA3_v(1).angle(n21%), A(0), 0, 0, A(1), 0, A(2), ty(0), 0, 1)
 Call combine_two_angle(tA3_v(0).angle(n12%), tA3_v(1).angle(n22%), A(3), 0, 0, A(4), 0, A(5), ty(1), 0, 1)
  tn(0) = 0
       ts(0) = tA3_v(0).para(n11%)
       tA3_v(0).para(0) = time_string(tA3_v(0).para(0), tA3_v(1).para(n21%), True, False)
       tA3_v(0).para(1) = time_string(tA3_v(0).para(1), tA3_v(1).para(n21%), True, False)
       tA3_v(0).para(2) = time_string(tA3_v(0).para(2), tA3_v(1).para(n21%), True, False)
       tA3_v(0).value = time_string(tA3_v(0).value, tA3_v(1).para(n21%), True, False)
       tA3_v(1).para(0) = time_string(tA3_v(1).para(0), ts(0), True, False)
       tA3_v(1).para(1) = time_string(tA3_v(1).para(1), ts(0), True, False)
       tA3_v(1).para(2) = time_string(tA3_v(1).para(2), ts(0), True, False)
       tA3_v(1).value = time_string(tA3_v(1).value, ts(0), True, False)
     If ty(0) = 3 Or ty(0) = 5 Then
       tA(0) = A(2)
     ElseIf ty(0) = 9 Or ty(0) = 10 Then
       tA(0) = A(2)
        tA3_v(0).para(n11%) = time_string(tA3_v(0).para(n11%), "-1", True, False)
        tA3_v(1).para(n21%) = time_string(tA3_v(1).para(n21%), "-1", True, False)
       tA3_v(0).value = add_string(tA3_v(0).value, time_string(tA3_v(0).para(n11%), "360", False, _
            False), True, False)
     ElseIf ty(0) = 4 Or ty(0) = 8 Then
       tA3_v(1).para(0) = time_string(tA3_v(1).para(0), "-1", True, False)
       tA3_v(1).para(1) = time_string(tA3_v(1).para(1), "-1", True, False)
       tA3_v(1).para(2) = time_string(tA3_v(1).para(2), "-1", True, False)
       tA3_v(1).value = time_string(tA3_v(1).value, "-1", True, False)
       If ty(0) = 4 Then
        tA(0) = A(0)
       Else
        tA(0) = A(1)
       End If
     ElseIf ty(0) = 6 Or ty(0) = 7 Then
       Call exchange_two_integer(tA3_v(0).angle(n11%), tA3_v(1).angle(n21%))
       Call exchange_two_integer(tA3_v(0).angle(n12%), tA3_v(1).angle(n22%))
       Call exchange_two_integer(tA3_v(0).angle(t_n(0)), tA3_v(1).angle(t_n(1)))
       Call exchange_two_string(tA3_v(0).para(n11%), tA3_v(1).para(n21%))
       Call exchange_two_string(tA3_v(0).para(n12%), tA3_v(1).para(n22%))
       Call exchange_two_string(tA3_v(0).para(t_n(0)), tA3_v(1).para(t_n(1)))
       Call exchange_two_string(tA3_v(0).value, tA3_v(1).value)
       tA3_v(1).para(0) = time_string(tA3_v(1).para(0), "-1", True, False)
       tA3_v(1).para(1) = time_string(tA3_v(1).para(1), "-1", True, False)
       tA3_v(1).para(2) = time_string(tA3_v(1).para(2), "-1", True, False)
       tA3_v(1).value = time_string(tA3_v(0).value, "-1", True, False)
       If ty(1) = 4 Then
        ty(1) = 6
       ElseIf ty(1) = 6 Then
        ty(1) = 4
       ElseIf ty(1) = 7 Then
        ty(1) = 8
       ElseIf ty(1) = 8 Then
        ty(1) = 7
       End If
       If ty(0) = 6 Then
        tA(0) = A(0)
       Else
        tA(0) = A(1)
       End If
     Else
       Exit Function
     End If
     '********************8
     If ty(1) = 3 Or ty(1) = 5 Then
        If tA3_v(0).para(n12%) = tA3_v(1).para(n22%) Then
         tA(1) = A(5)
        Else
         Exit Function
        End If
     ElseIf ty(1) = 9 Or ty(1) = 10 Then
        If tA3_v(0).para(n12%) = tA3_v(1).para(n22%) Then
         tA3_v(0).para(n12%) = time_string(tA3_v(0).para(n12%), "-1", True, False)
         tA3_v(1).para(n22) = time_string(tA3_v(1).para(n22%), "-1", True, False)
         tA(1) = A(5)
          tA3_v(0).value = add_string(tA3_v(0).value, time_string(tA3_v(0).para(n12%), "360", False, False), True, False)
        Else
         Exit Function
        End If
     ElseIf ty(1) = 4 Then
        If tA3_v(0).para(n12%) = time_string("-1", tA3_v(1).para(n22%), True, False) Then
         tA(1) = A(3)
        Else
         Exit Function
        End If
     ElseIf ty(1) = 6 Then
        If tA3_v(0).para(n12%) = time_string("-1", tA3_v(1).para(n22%), True, False) Then
         tA3_v(0).para(n12%) = tA3_v(1).para(n22%)
         tA(1) = A(3)
        Else
         Exit Function
        End If
     ElseIf ty(1) = 7 Then
        If tA3_v(0).para(n12%) = time_string("-1", tA3_v(1).para(n22%), True, False) Then
         tA3_v(0).para(n12%) = tA3_v(1).para(n22%)
         tA(1) = A(4)
        Else
         Exit Function
        End If
       ElseIf ty(1) = 8 Then
        If tA3_v(0).para(n12%) = time_string("-1", tA3_v(1).para(n22%), True, False) Then
         tA(1) = A(4)
        Else
         Exit Function
        End If
     Else
       Exit Function
     End If
     '*********************************
     t_s$ = add_string(tA3_v(0).value, tA3_v(1).value, True, False)
     If tA3_v(0).para(n12%) = "0" Then
        combine_three_angle_with_three_angle1 = set_three_angle_value( _
         tA(0), tA3_v(0).angle(t_n(0)), tA3_v(1).angle(t_n(1)), tA3_v(0).para(n11%), _
          tA3_v(0).para(t_n(0)), tA3_v(1).para(t_n(1)), t_s$, 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
       If combine_three_angle_with_three_angle1 Then
        Exit Function
       End If
     ElseIf is_two_angle_value(tA(0), tA(1), tA3_v(0).para(n11%), _
                 tA3_v(0).para(n12%), "", "", tn(0), tn(1), 0, 0, "", "", "") Then
            If angle3_value(tn(0)).data(0).data0.angle(0) = tA(0) Then
             ts(0) = tA3_v(0).para(n11%)
              ts(1) = tA3_v(0).para(n12%)
            ElseIf angle3_value(tn(0)).data(0).data0.angle(0) = tA(1) Then
             ts(0) = tA3_v(0).para(n12%)
              ts(1) = tA3_v(0).para(n11%)
            Else
             Exit Function
            End If
       If tn(1) = 0 Then
           t_s$ = minus_string(t_s$, time_string( _
             ts(0), angle3_value(tn(0)).data(0).data0.value, False, False), _
              True, False)
       Else
        t_s$ = minus_string(minus_string(t_s$, time_string( _
                  ts(0), angle3_value(tn(0)).data(0).data0.value, False, False), _
                   False, False), time_string( _
                  ts(1), angle3_value(tn(1)).data(0).data0.value, False, False), _
                   True, False)
       End If
       Call add_conditions_to_record(angle3_value_, tn(0), tn(1), 0, temp_record.record_data.data0.condition_data)
        combine_three_angle_with_three_angle1 = set_three_angle_value( _
         tA3_v(0).angle(t_n(0)), tA3_v(1).angle(t_n(1)), 0, tA3_v(0).para(t_n(0)), _
           tA3_v(1).para(t_n(1)), "0", t_s$, 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
      If combine_three_angle_with_three_angle1 Then
       Exit Function
      End If
    ElseIf is_two_angle_value(tA3_v(0).angle(t_n(0)), tA3_v(1).angle(t_n(1)), _
               tA3_v(0).para(t_n(0)), tA3_v(1).para(t_n(1)), "", "", tn(0), tn(1), 0, 0, "", "", "") Then
            If angle3_value(tn(0)).data(0).data0.angle(0) = tA3_v(0).angle(t_n(0)) Then
             ts(0) = tA3_v(0).para(t_n(0))
              ts(1) = tA3_v(0).para(t_n(1))
            ElseIf angle3_value(tn(0)).data(0).data0.angle(0) = tA3_v(1).angle(t_n(1)) Then
             ts(0) = tA3_v(1).para(t_n(1))
              ts(1) = tA3_v(0).para(t_n(0))
            Else
             Exit Function
            End If
      If tn(1) = 0 Then
        t_s$ = minus_string(t_s$, time_string( _
             ts(0), angle3_value(tn(0)).data(0).data0.value, False, False), _
              True, False)
      Else
        t_s$ = minus_string(minus_string(t_s$, time_string( _
                  ts(0), angle3_value(tn(0)).data(0).data0.value, False, False), _
                   False, False), time_string( _
                  ts(1), angle3_value(tn(1)).data(0).data0.value, False, False), _
                   True, False)
       End If
Call add_conditions_to_record(angle3_value_, tn(0), tn(1), 0, temp_record.record_data.data0.condition_data)
       combine_three_angle_with_three_angle1 = set_three_angle_value( _
        tA(0), tA(1), 0, tA3_v(0).para(n11%), tA3_v(0).para(n12%), "0", _
          t_s$, 0, temp_record, 0, 0, 0, no_reduce, 0, 0, False)
    If combine_three_angle_with_three_angle1 Then
     Exit Function
    End If
  End If
End Function

Public Function combine_three_line_with_eline(ByVal t%, _
          ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, no%, last_tn%
Dim n_(1) As Integer
Dim n(2) As Integer
Dim m(1) As Integer
Dim tn() As Integer
Dim el As eline_data0_type
 If line3_value(t%).record_.no_reduce > 4 Then
  Exit Function
 End If
 For i% = 0 To 2
 n(0) = i%
  n(1) = (i% + 1) Mod 3
   n(2) = (i% + 2) Mod 3
 For j% = 0 To 1
 m(0) = j
  m(1) = (j% + 1) Mod 2
 el.poi(2 * m(0)) = line3_value(t%).data(0).data0.poi(2 * n(0))
 el.poi(2 * m(0) + 1) = line3_value(t%).data(0).data0.poi(2 * n(0) + 1)
 el.poi(2 * m(1)) = -1
 Call search_for_eline(el, m(0), n_(0), 1)  '5.7
 el.poi(2 * m(1)) = 30000
 Call search_for_eline(el, m(0), n_(1), 1)
 last_tn% = 0
 For k% = n_(0) + 1 To n_(1)
 no% = Deline(k%).data(0).record.data1.index.i(m(0))
 If no% > start% And Deline(no%).record_.no_reduce < 4 Then
 'If is_two_record_related(eline_, no%, Deline(no%).data(0).record, _
      line3_value_, t%, line3_value(t%).data(0).record) = False Then
 last_tn% = last_tn% + 1
 ReDim Preserve tn(last_tn%) As Integer
 tn(last_tn%) = no%
 End If
 'End If
 Next k%
 For k% = 1 To last_tn%
 no% = tn(k%)
  combine_three_line_with_eline = _
   combine_relation_with_three_line_(eline_, no%, _
    line3_value_, t%, m(0), n(0))
   If combine_three_line_with_eline > 1 Then
   Exit Function
   End If
 Next k%
 Next j%
 Next i%
End Function

Public Function combine_three_line_with_midpoint(ByVal t%, _
          ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, no%
Dim v(1) As String
Dim n(2) As Integer
Dim m(5) As Integer
Dim md As mid_point_data0_type
If line3_value(t%).record_.no_reduce > 4 Then
 Exit Function
End If
For i% = 0 To 2
 n(0) = i%
  n(1) = (i% + 1) Mod 3
   n(2) = (i% + 2) Mod 3
For j% = 0 To 2
 If j% = 0 Then
  m(0) = 0
   m(1) = 1
    m(2) = 1
     m(3) = 2
      m(4) = 0
       m(5) = 2
 ElseIf j% = 1 Then
  m(0) = 1
   m(1) = 2
    m(2) = 0
     m(3) = 2
      m(4) = 0
       m(5) = 1
 Else
  m(0) = 0
   m(1) = 2
    m(2) = 0
     m(3) = 1
      m(4) = 1
       m(5) = 2
 End If
 md.poi(m(0)) = line3_value(t%).data(0).data0.poi(2 * n(0))
  md.poi(m(1)) = line3_value(t%).data(0).data0.poi(2 * n(0) + 1)
If search_for_mid_point(md, j%, no%, 2) Then  '5.7原j%+3
  If no% > start% And Dmid_point(no%).record_.no_reduce < 4 Then
   'If is_two_record_related(midpoint_, no%, Dmid_point(no%).data(0).record, _
         line3_value_, t%, line3_value(t%).data(0).record) = False Then
  combine_three_line_with_midpoint = _
    combine_relation_with_three_line_(midpoint_, no%, line3_value_, _
     t%, j%, n(0))
If combine_three_line_with_midpoint > 1 Then
   Exit Function
End If
End If
End If
'End If
Next j%
 Next i%
End Function

Public Function combine_three_line_with_line_value(ByVal t%, _
            ByVal start%, ByVal no_reduce As Byte) As Byte
Dim j%, l%
Dim n(2) As Integer
If line3_value(t%).record_.no_reduce > 4 Then
 Exit Function
End If
 For j% = 0 To 2
 n(0) = j%
 n(1) = (j% + 1) Mod 3
 n(2) = (j% + 2) Mod 3
  If is_line_value(line3_value(t%).data(0).data0.poi(2 * n(0)), _
      line3_value(t%).data(0).data0.poi(2 * n(0) + 1), _
       line3_value(t%).data(0).data0.n(2 * n(0)), _
      line3_value(t%).data(0).data0.n(2 * n(0) + 1), _
       line3_value(t%).data(0).data0.line_no(n(0)), _
         "", l%, -1000, _
       0, 0, 0, line_value_data0) = 1 Then
 If l% > start% Then
  If line_value(l%).record_.no_reduce < 255 Then
  combine_three_line_with_line_value = _
    combine_three_three_line_(line_value_, l%, _
     line3_value_, t%, 0, n(0), 0)
If combine_three_line_with_line_value > 1 Then
   Exit Function
  End If
End If
End If
End If
 Next j%
For j% = 1 To last_conditions.last_cond(1).line_value_no
 If InStr(1, line_value(j%).data(0).data0.value_, "x", 0) > 0 Then
  combine_three_line_with_line_value = subs_line_value_to_line3_value(t%, j%)
  If combine_three_line_with_line_value > 1 Then
      Exit Function
  End If
 End If
Next j%
End Function

Public Function combine_two_line0(l_data1 As line_data_type, l_data2 As line_data_type) As line_data_type
Dim i%, st%
Dim t_line_data0 As line_data_type
t_line_data0 = l_data1
For i% = l_data2.data0.in_point(0) To 1 Step -1
combine_two_line0 = add_point_to_line_data(l_data2.data0.in_point(i%), t_line_data0, 0, False, True, st%)
t_line_data0 = combine_two_line0
Next i%
If l_data2.is_change Then
   combine_two_line0.is_change = True
End If
End Function
Public Function combine_two_line(ByVal l1%, ByVal l2%, ByVal link_p%, _
                   re As record_data_type, no_reduce As Byte, _
                        ByVal is_no_initial As Byte) As Byte
Dim i%, j%, k%, l3%, tl%, no%, tn_%, new_p%, p3_no%, st%
Dim tn(8) As Integer
Dim tp(1) As Integer
Dim tl_(10) As Integer
Dim last_tl%
Dim t As Boolean
Dim T_lin(2) As line_data_type
Dim l_data0 As line_data_type
Dim temp_record As total_record_type
Dim ty As Byte
'On Error GoTo combine_two_line_error:
Call set_level(re.data0.condition_data)
If l1% = l2% Or l1% = 0 Or l2% = 0 Then
Exit Function
End If
For i% = 1 To last_conditions.last_cond(1).same_three_lines_no
 l3% = same_three_lines(i%).data(0).line_no(2)
  tl_(0) = l1%
   tl_(1) = l2%
  If l1% = l3% Then
     tl_(0) = 0
  Else
  'If m_lin(l1%).data(0).no_reduce Then
   If m_lin(l1%).data(0).other_no = l3% Then
     tl_(0) = 0
   End If
  'End If
  End If
  If l2% = l3% Then
  Else
  'If m_lin(l2%).data(0).no_reduce Then
   If m_lin(l2%).data(0).other_no = l3% Then
     tl_(1) = 0
   End If
  'End If
  End If
  For j% = 2 To m_lin(l3%).data(0).data0.in_point(0)
  'l1%和l2%与某合并线重合
   For k% = 1 To j% - 1
        If tl_(0) <> 0 Then
         If m_lin(l1%).data(0).data0.poi(0) = m_lin(l3%).data(0).data0.in_point(k%) And _
            m_lin(l1%).data(0).data0.poi(1) = m_lin(l3%).data(0).data0.in_point(j%) Then
              tl_(0) = 0
         End If
        End If
        If tl_(1) <> 0 Then
         If m_lin(l2%).data(0).data0.poi(0) = m_lin(l3%).data(0).data0.in_point(k%) And _
            m_lin(l2%).data(0).data0.poi(1) = m_lin(l3%).data(0).data0.in_point(j%) Then
              tl_(1) = 0
         End If
        End If
   Next k%
  Next j%
  If tl_(0) = 0 And tl_(1) = 0 Then
    Exit Function '全重合
  ElseIf tl_(0) = 0 Then
    GoTo combine_two_line_mark0
  ElseIf tl_(1) = 0 Then
    GoTo combine_two_line_mark0
  End If
Next i%
'退化直线
For i% = 1 To last_conditions.last_cond(1).line_no
 'If m_lin(i%).data(0).no_reduce = True Then
    If m_lin(i%).data(0).other_no = l1% Or _
          m_lin(i%).data(0).other_no = l2% Then
          m_lin(i%).data(0).other_no = l3%
    End If
 'End If
Next i%
If last_conditions.last_cond(1).same_three_lines_no Mod 10 = 0 Then
ReDim Preserve same_three_lines(last_conditions.last_cond(1).same_three_lines_no + 10) _
    As same_three_lines_type
End If
last_conditions.last_cond(1).same_three_lines_no = last_conditions.last_cond(1).same_three_lines_no + 1
no% = last_conditions.last_cond(1).same_three_lines_no
same_three_lines(last_conditions.last_cond(1).same_three_lines_no).data(0).record = re
If m_lin(l1%).data(0).data0.poi(0) = m_lin(l2%).data(0).data0.poi(0) Then 'l3% 是长线
 If compare_two_point(m_poi(m_lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate, _
                  m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate, _
                   m_lin(l1%).data(0).data0.poi(0), m_lin(l1%).data(0).data0.poi(1), 6) = 1 Then
   l3% = l2%
    l2% = line_number0(m_lin(l1%).data(0).data0.poi(1), m_lin(l2%).data(0).data0.poi(1), tn(0), tn(1))
     new_p% = m_lin(l1%).data(0).data0.poi(1)
      tl% = l2%
 Else
   l3% = l1%
    l1% = line_number0(m_lin(l2%).data(0).data0.poi(1), m_lin(l1%).data(0).data0.poi(1), tn(0), tn(1))
      new_p% = m_lin(l2%).data(0).data0.poi(1)
    tl% = l1%
 End If
' tl% = l3%
ElseIf m_lin(l1%).data(0).data0.poi(0) = m_lin(l2%).data(0).data0.poi(1) Then
  new_p% = m_lin(l1%).data(0).data0.poi(0)
 l3% = line_number0(m_lin(l1%).data(0).data0.poi(1), m_lin(l2%).data(0).data0.poi(0), tn(0), tn(1))
  tl% = l3%
ElseIf m_lin(l1%).data(0).data0.poi(1) = m_lin(l2%).data(0).data0.poi(0) Then
  new_p% = m_lin(l1%).data(0).data0.poi(1)
 l3% = line_number0(m_lin(l1%).data(0).data0.poi(0), m_lin(l2%).data(0).data0.poi(1), tn(0), tn(1))
  tl% = l3%
ElseIf m_lin(l1%).data(0).data0.poi(1) = m_lin(l2%).data(0).data0.poi(1) Then
 If compare_two_point(m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                 m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                   m_lin(l1%).data(0).data0.poi(0), m_lin(l1%).data(0).data0.poi(1), 6) = 1 Then
   l3% = l1%
    l1% = line_number0(m_lin(l2%).data(0).data0.poi(0), m_lin(l1%).data(0).data0.poi(0), tn(0), tn(1))
      new_p% = m_lin(l2%).data(0).data0.poi(0)
       tl% = l3%
 Else
   l3% = l2%
    l2% = line_number0(m_lin(l1%).data(0).data0.poi(0), m_lin(l2%).data(0).data0.poi(0), tn(0), tn(1))
     new_p% = m_lin(l1%).data(0).data0.poi(0)
      tl% = l3%
 End If
' tl% = l3%
Else
 If compare_two_point(m_poi(m_lin(l1%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                       m_poi(m_lin(l2%).data(0).data0.poi(0)).data(0).data0.coordinate, _
                         m_lin(l1%).data(0).data0.poi(0), m_lin(l1%).data(0).data0.poi(1), 6) = 1 Then
  tp(0) = m_lin(l1%).data(0).data0.poi(0)
 Else
  tp(0) = m_lin(l2%).data(0).data0.poi(0)
 End If
 If compare_two_point(m_poi(m_lin(l1%).data(0).data0.poi(1)).data(0).data0.coordinate, _
                       m_poi(m_lin(l2%).data(0).data0.poi(1)).data(0).data0.coordinate, _
                        m_lin(l1%).data(0).data0.poi(0), m_lin(l1%).data(0).data0.poi(1), 6) = 1 Then
  tp(1) = m_lin(l2%).data(0).data0.poi(1)
 Else
  tp(1) = m_lin(l1%).data(0).data0.poi(1)
 End If
 l3% = line_number0(tp(0), tp(1), tn(0), tn(1))
 tl% = l3%
End If
For i% = 1 To last_conditions.last_cond(1).line_no
If i% <> l1% And i% <> l2% And i% <> l3% Then
   tp(0) = is_line_line_intersect(i%, l1%, 0, 0, False)
   tp(1) = is_line_line_intersect(i%, l2%, 0, 0, False)
 If tp(0) <> 0 And tp(1) <> 0 And tp(0) <> tp(1) Then
 temp_record.record_data = re
  combine_two_line = combine_two_point(tp(0), tp(1), 0, temp_record)
  If combine_two_line > 1 Then
     Exit Function
  End If
 End If
 End If
Next i%
 If l1% < l2% Then
 same_three_lines(no%).data(0).line_no(0) = l1%
 same_three_lines(no%).data(0).line_no(1) = l2%
 Else 'If l1% < l2% Then
 same_three_lines(no%).data(0).line_no(0) = l2%
 same_three_lines(no%).data(0).line_no(1) = l1%
 End If
 same_three_lines(no%).data(0).line_no(2) = l3%
 'Call set_line_no_reduce(l1%, True)
 Call set_line_other_no(l1%, l3%)
 'Call set_line_no_reduce(l2%, True)
 Call set_line_other_no(l2%, l3%)
 'lin(l2%).data(0).no_reduce = True
 'lin(l2%).data(0).other_no = l3%
 Call set_line_cond_data(l3%, re.data0.condition_data)
For i% = 1 To m_lin(l3%).data(0).data0.in_point(0)
    tl% = line_number0(m_lin(l1%).data(0).data0.poi(0), m_lin(l3%).data(0).data0.in_point(i%), 0, 0)
     If tl% <> l3% And tl% > 0 Then
      For j% = 1 To last_tl%
       If tl_(j%) = tl% Then
          GoTo combine_two_line_mark12
       End If
      Next j%
      last_tl% = last_tl% + 1
      tl_(last_tl%) = tl%
     End If
combine_two_line_mark12:
    tl% = line_number0(m_lin(l1%).data(0).data0.poi(1), m_lin(l3%).data(0).data0.in_point(i%), 0, 0)
     If tl% <> l3% And tl% > 0 Then
      For j% = 1 To last_tl%
       If tl_(j%) = tl% Then
          GoTo combine_two_line_mark13
       End If
      Next j%
      last_tl% = last_tl% + 1
      tl_(last_tl%) = tl%
     End If
combine_two_line_mark13:
    tl% = line_number0(m_lin(l2%).data(0).data0.poi(0), m_lin(l3%).data(0).data0.in_point(i%), 0, 0)
     If tl% <> l3% And tl% > 0 Then
      For j% = 1 To last_tl%
       If tl_(j%) = tl% Then
          GoTo combine_two_line_mark14
       End If
      Next j%
      last_tl% = last_tl% + 1
      tl_(last_tl%) = tl%
     End If
combine_two_line_mark14:
    tl% = line_number0(m_lin(l2%).data(0).data0.poi(1), m_lin(l3%).data(0).data0.in_point(i%), 0, 0)
     If tl% <> l3% And tl% > 0 Then
      For j% = 1 To last_tl%
       If tl_(j%) = tl% Then
          GoTo combine_two_line_mark15
       End If
      Next j%
      last_tl% = last_tl% + 1
      tl_(last_tl%) = tl%
     End If
combine_two_line_mark15:
Next i%
For i% = 1 To last_tl%
 'Call set_line_no_reduce(tl_(i%), True)
 Call set_line_other_no(tl_(i%), l3%)
Next i%
For i% = m_lin(l1%).data(0).data0.in_point(0) To 1 Step -1
 combine_two_line = add_point_to_line(m_lin(l1%).data(0).data0.in_point(i%), l3%, _
                0, False, False, st%)
                If combine_two_line > 1 Then
                   Exit Function
                End If
Next i%
st% = 0
For k% = m_lin(l2%).data(0).data0.in_point(0) To 1 Step -1 ' To m_lin(l2%).data(0).data0.in_point(0)
    combine_two_line = add_point_to_line(m_lin(l2%).data(0).data0.in_point(k%), l3%, _
                0, False, False, st%)
                If combine_two_line > 1 Then
                   Exit Function
                End If
Next k%
For i% = 2 To m_lin(l3%).data(0).data0.in_point(0)
 For j% = 1 To i% - 1
  If search_for_two_point_line(m_lin(l3%).data(0).data0.in_point(i%), m_lin(l3%).data(0).data0.in_point(j%), _
                   tn_%, 0) Then
    'If tl% <> l3% Then
     Dtwo_point_line(tn_%).data(0).line_no = l3%
     If m_lin(l3%).data(0).data0.in_point(i%) < m_lin(l3%).data(0).data0.in_point(j%) Then
     Dtwo_point_line(tn_%).data(0).n(0) = i%
     Dtwo_point_line(tn_%).data(0).n(1) = j%
     Else
     Dtwo_point_line(tn_%).data(0).n(0) = i%
     Dtwo_point_line(tn_%).data(0).n(1) = j%
     End If
     If regist_data.run_type = 1 Then
        If m_lin(l3%).data(0).data0.in_point(10) = 0 Then
           m_lin(l3%).data(0).data0.in_point(10) = 1
        End If
        Dtwo_point_line(tn_%).data(0).dir = m_lin(l3%).data(0).data0.in_point(10)
        If m_lin(l3%).data(0).data0.in_point(10) = 1 Then
         Dtwo_point_line(tn_%).data(0).v_poi(0) = Dtwo_point_line(tn_%).poi(0)
         Dtwo_point_line(tn_%).data(0).v_poi(1) = Dtwo_point_line(tn_%).poi(1)
         Dtwo_point_line(tn_%).data(0).v_n(0) = Dtwo_point_line(tn_%).data(0).n(0)
         Dtwo_point_line(tn_%).data(0).v_n(1) = Dtwo_point_line(tn_%).data(0).n(1)
        Else
         Dtwo_point_line(tn_%).data(0).v_poi(0) = Dtwo_point_line(tn_%).poi(1)
         Dtwo_point_line(tn_%).data(0).v_poi(1) = Dtwo_point_line(tn_%).poi(0)
         Dtwo_point_line(tn_%).data(0).v_n(0) = Dtwo_point_line(tn_%).data(0).n(1)
         Dtwo_point_line(tn_%).data(0).v_n(1) = Dtwo_point_line(tn_%).data(0).n(0)
        End If
     End If
   End If
  For k% = 1 To j% - 1
   temp_record.record_data = re
           no% = 0
           combine_two_line = _
             max_for_byte(combine_two_line, _
              set_three_point_on_line(m_lin(l3%).data(0).data0.in_point(i%), m_lin(l3%).data(0).data0.in_point(j%), _
              m_lin(l3%).data(0).data0.in_point(k%), temp_record, no%, 0, is_no_initial))
          If k% = 1 And i% = m_lin(l3%).data(0).data0.in_point(0) Then
             If m_lin(l3%).data(0).data0.in_point(j%) = m_lin(l1%).data(0).data0.poi(0) Or _
                m_lin(l3%).data(0).data0.in_point(j%) = m_lin(l1%).data(0).data0.poi(1) Or _
                m_lin(l3%).data(0).data0.in_point(j%) = m_lin(l2%).data(0).data0.poi(0) Or _
                m_lin(l3%).data(0).data0.in_point(j%) = m_lin(l2%).data(0).data0.poi(1) Then
                  p3_no% = no%
             End If
          End If
          If combine_two_line > 1 Then
             Exit Function
          End If
  Next k%
 Next j%
Next i%
'****************************************************
combine_two_line_mark0:
temp_record.record_data.data0.condition_data.condition_no = 1
temp_record.record_data.data0.condition_data.condition(1).ty = point3_on_line_
temp_record.record_data.data0.condition_data.condition(1).no = p3_no%
For i% = 0 To last_tl% \ 2
 If tl_(2 * i% + 1) > 0 Or tl_(2 * i% + 2) > 0 Then
 combine_two_line = simple_dbase_for_line(l3%, tl_(2 * i% + 1), tl_(2 * i% + 2), new_p%, temp_record.record_data)
 End If
Next i%
'Call set_two_point_line_for_line(l3%, re)
combine_two_line = set_total_equal_triangle_from_combine_two_line(l3%, new_p%, temp_record)
If combine_two_line > 1 Then
Exit Function
End If
combine_two_line = set_total_equal_triangle_from_combine_two_line(l3%, new_p%, temp_record)
combine_two_line_error:
combine_two_line = 0
'************************************
End Function
Public Function combine_two_line_with_eline(ByVal t%, _
          ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, no%, last_tn%
Dim n(1) As Integer
Dim m(1) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim el As eline_data0_type
If two_line_value(t%).record_.no_reduce > 4 Then
 Exit Function
End If
For i% = 0 To 1
 n(0) = i%
  n(1) = (i% + 1) Mod 2
For j% = 0 To 1
   m(0) = j%
    m(1) = (j% + 1) Mod 2
el.poi(2 * m(0)) = two_line_value(t%).data(0).data0.poi(2 * n(0))
el.poi(2 * m(0) + 1) = two_line_value(t%).data(0).data0.poi(2 * n(0) + 1)
el.poi(2 * m(1)) = 0
Call search_for_eline(el, m(0), n_(0), 1)  '5.7
el.poi(2 * m(1)) = 30000
Call search_for_eline(el, m(0), n_(1), 1)
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
no% = Deline(k%).data(0).record.data1.index.i(m(0))
If no% > start% And Deline(no%).record_.no_reduce < 4 Then
'If is_two_record_related(eline_, no%, Deline(no%).data(0).record, _
      two_line_value_, t%, two_line_value(t%).data(0).record) = False Then
last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
End If
'End If
Next k%
For k% = 1 To last_tn%
no% = tn(k%)
combine_two_line_with_eline = combine_relation_with_three_line_( _
  eline_, no%, two_line_value_, t%, m(0), n(0))
If combine_two_line_with_eline > 1 Then
  Exit Function
End If
Next k%
Next j%
Next i%
End Function

Public Function combine_two_line_with_mid_point(ByVal t%, _
            ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, no%
Dim n(1) As Integer
Dim m(5) As Integer
Dim md As mid_point_data0_type
Dim s(1) As String
For i% = 0 To 1
 n(0) = i%
  n(1) = (i% + 1) Mod 2
 For j% = 0 To 2
  If j% = 0 Then
   m(0) = 0
   m(1) = 1
   m(2) = 1
   m(3) = 2
   m(4) = 0
   m(5) = 2
 ElseIf j% = 1 Then
   m(0) = 1
   m(1) = 2
   m(2) = 0
   m(3) = 2
   m(4) = 1
   m(5) = 2
 Else
   m(0) = 0
   m(1) = 2
   m(2) = 0
   m(3) = 1
   m(4) = 1
   m(5) = 2
End If
md.poi(m(0)) = two_line_value(t%).data(0).data0.poi(2 * n(0))
md.poi(m(1)) = two_line_value(t%).data(0).data0.poi(2 * n(0) + 1)
If search_for_mid_point(md, j%, no%, 2) Then  '5.7原j%+3
If no% > start% And Dmid_point(no%).record_.no_reduce < 4 Then
'If is_two_record_related(midpoint_, no%, Dmid_point(no%).data(0).record, _
       two_line_value_, t%, two_line_value(t%).data(0).record) = False Then
 combine_two_line_with_mid_point = _
    combine_relation_with_three_line_(midpoint_, no%, two_line_value_, _
     t%, j%, n(0))
  If combine_two_line_with_mid_point > 1 Then
   Exit Function
  End If
 End If
 End If
 'End If
 Next j%
 Next i%
End Function

Public Function combine_two_line_with_line_value(ByVal t%, _
          ByVal start%, ByVal no_reduce As Byte) As Byte
Dim j%, no%
Dim n(1) As Integer
Dim temp_record As total_record_type
Dim re As record_data_type
If two_line_value(t%).record_.no_reduce > 4 Then
 Exit Function
End If
re.data0.condition_data.condition_no = 0 ' record0
re.data0.theorem_no = 1
Call add_conditions_to_record(two_line_value_, t%, 0, 0, re.data0.condition_data)
 For j% = 0 To 1
 n(0) = j%
  n(1) = (j% + 1) Mod 2
  If is_line_value(two_line_value(t%).data(0).data0.poi(2 * n(0)), _
       two_line_value(t%).data(0).data0.poi(2 * n(0) + 1), _
        two_line_value(t%).data(0).data0.n(2 * n(0)), _
         two_line_value(t%).data(0).data0.n(2 * n(0) + 1), _
          two_line_value(t%).data(0).data0.line_no(n(0)), "", no%, _
        -1000, 0, 0, 0, line_value_data0) = 1 Then
   If no% > start% Then
   If line_value(no%).record_.no_reduce < 255 Then
   'If is_two_record_related(line_value_, no%, line_value(no%).data(0).record, _
       two_line_value_, t%, two_line_value(t%).data(0).record) = False Then
    temp_record.record_data = re
    Call add_conditions_to_record(line_value_, no%, 0, 0, temp_record.record_data.data0.condition_data)
  combine_two_line_with_line_value = _
    set_line_value(two_line_value(t%).data(0).data0.poi(2 * n(1)), _
          two_line_value(t%).data(0).data0.poi(2 * n(1) + 1), _
           divide_string(minus_string(two_line_value(t%).data(0).data0.value, _
            time_string(line_value(no%).data(0).data0.value, two_line_value(t%).data(0).data0.para(n(0)), _
             False, False), False, False), _
             two_line_value(t%).data(0).data0.para(n(1)), True, False), two_line_value(t%).data(0).data0.n(2 * n(1)), _
              two_line_value(t%).data(0).data0.n(2 * n(1) + 1), two_line_value(t%).data(0).data0.line_no(n(1)), _
               temp_record, 0, 0, False)
If combine_two_line_with_line_value > 1 Then
  Exit Function
End If
End If
 End If
 End If
' End If
 Next j%
For j% = 1 To last_conditions.last_cond(1).line_value_no
 If InStr(1, line_value(j%).data(0).data0.value_, "x", 0) > 0 Then
  combine_two_line_with_line_value = subs_line_value_to_two_line_value(t%, j%)
  If combine_two_line_with_line_value > 1 Then
      Exit Function
  End If
 End If
Next j%
End Function

Public Function combine_two_line_with_relation(ByVal t%, _
         ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, no%, last_tn%
Dim n(1) As Integer
Dim m(2) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim v(1) As String
Dim re As relation_data0_type
For i% = 0 To 1
 n(0) = i%
  n(1) = (i% + 1) Mod 2
 For j% = 0 To 2
  m(0) = j%
   m(1) = (j% + 1) Mod 3
    m(2) = (j% + 2) Mod 3
 re.poi(2 * m(0)) = two_line_value(t%).data(0).data0.poi(2 * n(0))
 re.poi(2 * m(0) + 1) = two_line_value(t%).data(0).data0.poi(2 * n(0) + 1)
 re.poi(2 * m(1)) = -1
 Call search_for_relation(re, m(0), n_(0), 1)  '5.7
 re.poi(2 * m(1)) = 30000
 Call search_for_relation(re, m(0), n_(1), 1)
 last_tn% = 0
 For k% = n_(0) + 1 To n_(1)
 no% = Drelation(k%).data(0).record.data1.index.i(m(0))
 If no% > start% And Drelation(no%).record_.no_reduce < 4 Then
  'If is_two_record_related(relation_, no%, Drelation(no%).data(0).record, _
          two_line_value_, t%, two_line_value(t%).data(0).record) = False Then
 last_tn% = last_tn% + 1
  ReDim Preserve tn(last_tn%) As Integer
   tn(last_tn%) = no%
 End If
 'End If
 Next k%
 For k% = 1 To last_tn%
  no% = tn(k%)
   combine_two_line_with_relation = combine_relation_with_three_line_( _
    relation_, no%, two_line_value_, t%, m(0), n(0))
   If combine_two_line_with_relation > 1 Then
    Exit Function
   End If
 Next k%
Next j%
Next i%
End Function
Public Function combine_three_three_line_(ByVal ty1 As Byte, ByVal t1%, _
    ByVal ty2 As Byte, ByVal t2%, ByVal k%, ByVal l%, no_reduce As Byte) As Byte
Dim i%, j%, m%
Dim p1(7) As Integer
Dim n1(7) As Integer
Dim l1(3) As Integer
Dim p2(5) As Integer
Dim n2(5) As Integer
Dim l2(2) As Integer
Dim para1(3) As String
Dim para2(2) As String
Dim v1$
Dim v2$
Dim temp_record As total_record_type
temp_record.record_data.data0.condition_data.condition_no = 0 ' record0
Call add_conditions_to_record(ty1, t1%, 0, 0, temp_record.record_data.data0.condition_data)
Call add_conditions_to_record(ty2, t2%, 0, 0, temp_record.record_data.data0.condition_data)
temp_record.record_data.data0.theorem_no = 1
Call read_point_and_value_from_line_value(ty1, t1%, k%, p1(), n1(), _
        l1(), para1(), v1$)
Call read_point_and_value_from_line_value(ty2, t2%, l%, p2(), n2(), _
        l2(), para2(), v2$)
para2(1) = time_string(para2(1), para1(0), True, False)
para2(2) = time_string(para2(2), para1(0), True, False)
v2$ = time_string(v2$, para1(0), True, False)
para1(1) = time_string(para1(1), para2(0), True, False)
para1(2) = time_string(para1(2), para2(0), True, False)
para1(0) = time_string("-1", para2(1), True, False)
para1(3) = time_string("-1", para2(2), True, False)
v1$ = time_string(v1$, para2(0), True, False)
p1(0) = p2(2)
p1(1) = p2(3)
p1(6) = p2(4)
p1(7) = p2(5)
n1(0) = n2(2)
n1(1) = n2(3)
n1(6) = n2(4)
n1(7) = n2(5)
l1(0) = l2(1)
l1(3) = l2(2)
v1$ = minus_string(v1$, v2$, True, False)
For i% = 1 To 3
 For j% = 0 To i% - 1
  If p1(2 * i%) = p1(2 * j%) And p1(2 * i% + 1) = p1(2 * j% + 1) Then
   p1(2 * i%) = 0
   p1(2 * i% + 1) = 0
   n1(2 * i%) = 0
   n1(2 * i% + 1) = 0
   l1(i%) = 0
   para1(j%) = add_string(para1(i%), para1(j%), True, False)
   If para1(j%) = "0" Then
    p1(2 * j%) = 0
    p1(2 * j% + 1) = 0
    l1(j%) = 0
   End If
   para1(i%) = "0"
  End If
 Next j%
Next i%
If para1(0) = "0" And para1(1) = "0" And para1(2) = "0" And para1(3) = "0" Then
 Exit Function
End If
For m% = 0 To 2
For i% = 0 To 3
 If para1(i%) = "0" Then
  For j% = i% To 2
   p1(2 * j%) = p1(2 * j% + 2)
    p1(2 * j% + 1) = p1(2 * j% + 3)
   n1(2 * j%) = n1(2 * j% + 2)
    n1(2 * j% + 1) = n1(2 * j% + 3)
   l1(j%) = l1(j% + 1)
     para1(j%) = para1(j% + 1)
  Next j%
  p1(6) = 0
   p1(7) = 0
  n1(6) = 0
   n1(7) = 0
  l1(3) = 0
    para1(3) = "0"
 End If
Next i%
Next m%
If para1(3) = "0" Then
 combine_three_three_line_ = set_three_line_value( _
  p1(0), p1(1), p1(2), p1(3), p1(4), p1(5), n1(0), n1(1), n1(2), _
     n1(3), n1(4), n1(5), l1(0), l1(1), l1(2), _
   para1(0), para1(1), para1(2), v1$, temp_record, 0, no_reduce, 0)
If ty1 = line_value_ Then
Call set_record_no_reduce(ty2, t2%, 0, 0, 255)
ElseIf ty2 = line_value_ Then
Call set_record_no_reduce(ty1, t2%, 0, 0, 255)
ElseIf ty1 = two_line_value_ Then
Call set_record_no_reduce(ty2, t2%, 0, 0, 255)
ElseIf ty2 = two_line_value_ Then
Call set_record_no_reduce(ty1, t2%, 0, 0, 255)
Else
Call set_record_no_reduce(ty2, t2%, 0, 0, 255)
End If
End If
End Function
Public Function combine_three_three_line(ByVal t1%, _
    ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, no%, l%, p%, o%
Dim last_tn11%, last_tn12%, last_tn21%, last_tn22%
Dim n(2) As Integer
Dim m(2) As Integer
Dim n_(1) As Integer
Dim tn11() As Integer
Dim tn12() As Integer
Dim tn21() As Integer
Dim tn22() As Integer
Dim tp(7) As Integer
Dim tn(7) As Integer
Dim tl(3) As Integer
Dim s(2) As String
Dim v As String
Dim t_l As line3_value_data0_type
Dim temp_record As total_record_type
Dim re As record_data_type
If line3_value(t1%).record_.no_reduce > 4 Then
     Exit Function
End If
re.data0.theorem_no = 1
Call add_conditions_to_record(line3_value_, t1%, 0, 0, re.data0.condition_data)
For i% = 0 To 3
 If i% < 2 Then
 n(0) = i%
  n(1) = (i% + 1) Mod 3
   n(2) = (i% + 2) Mod 3
 ElseIf i% = 3 Then
 n(0) = 0
  n(1) = 2
   n(2) = 1
 Else
 GoTo combine_three_three_line_next1
 End If
For j% = 0 To 3
If j% < 2 Then
 m(0) = j%
  m(1) = (j% + 1) Mod 3
   m(2) = (j% + 2) Mod 3
ElseIf j% = 3 Then
 m(0) = 0
  m(1) = 2
   m(2) = 1
Else
 GoTo combine_three_three_line_next2
End If
t_l.poi(2 * m(0)) = line3_value(t1%).data(0).data0.poi(2 * n(0))
 t_l.poi(2 * m(0) + 1) = line3_value(t1%).data(0).data0.poi(2 * n(0) + 1)
  t_l.poi(2 * m(1)) = -1
Call search_for_line3_value(t_l, m(0), n_(0), 1)
t_l.poi(2 * m(1)) = 30000
Call search_for_line3_value(t_l, m(0), n_(1), 1)  '5.7
last_tn11% = 0
last_tn12% = 0
last_tn21% = 0
last_tn22% = 0
For k% = n_(0) + 1 To n_(1)
no% = line3_value(k%).data(0).record.data1.index.i(m(0))
If no% > 0 And no% < t1% And _
  line3_value(no%).record_.no_reduce < 4 Then
'If is_two_record_related(line3_value_, no%, line3_value(no%).data(0).record, _
       line3_value_, t1%, line3_value(t1%).data(0).record) = False Then
If line3_value(no%).data(0).data0.poi(2 * m(1)) = line3_value(t1%).data(0).data0.poi(2 * n(1)) And _
    line3_value(no%).data(0).data0.poi(2 * m(1) + 1) = line3_value(t1%).data(0).data0.poi(2 * n(1) + 1) Then
last_tn11% = last_tn11% + 1
ReDim Preserve tn11(last_tn11%) As Integer
tn11(last_tn11%) = no%
ElseIf line3_value(no%).data(0).data0.poi(2 * m(1)) = line3_value(t1%).data(0).data0.poi(2 * n(2)) And _
    line3_value(no%).data(0).data0.poi(2 * m(1) + 1) = line3_value(t1%).data(0).data0.poi(2 * n(2) + 1) Then
last_tn12% = last_tn12% + 1
ReDim Preserve tn12(last_tn12%) As Integer
tn12(last_tn12%) = no%
ElseIf line3_value(no%).data(0).data0.poi(2 * m(2)) = line3_value(t1%).data(0).data0.poi(2 * n(1)) And _
    line3_value(no%).data(0).data0.poi(2 * m(2) + 1) = line3_value(t1%).data(0).data0.poi(2 * n(1) + 1) Then
last_tn21% = last_tn21% + 1
ReDim Preserve tn21(last_tn21%) As Integer
tn21(last_tn21%) = no%
ElseIf line3_value(no%).data(0).data0.poi(2 * m(2)) = line3_value(t1%).data(0).data0.poi(2 * n(2)) And _
    line3_value(no%).data(0).data0.poi(2 * m(2) + 1) = line3_value(t1%).data(0).data0.poi(2 * n(2) + 1) Then
last_tn22% = last_tn22% + 1
ReDim Preserve tn22(last_tn22%) As Integer
tn22(last_tn22%) = no%
'End If
End If
End If
Next k%
For k% = 1 To last_tn11%
no% = tn11(k%)
temp_record.record_data = re
Call add_conditions_to_record(line3_value_, no%, 0, 0, temp_record.record_data.data0.condition_data)
  tp(0) = line3_value(t1%).data(0).data0.poi(2 * n(1))
  tp(1) = line3_value(t1%).data(0).data0.poi(2 * n(1) + 1)
  tp(2) = line3_value(t1%).data(0).data0.poi(2 * n(2))
  tp(3) = line3_value(t1%).data(0).data0.poi(2 * n(2) + 1)
  tp(4) = line3_value(no%).data(0).data0.poi(2 * m(2))
  tp(5) = line3_value(no%).data(0).data0.poi(2 * m(2) + 1)
  tn(0) = line3_value(t1%).data(0).data0.n(2 * n(1))
  tn(1) = line3_value(t1%).data(0).data0.n(2 * n(1) + 1)
  tn(2) = line3_value(t1%).data(0).data0.n(2 * n(2))
  tn(3) = line3_value(t1%).data(0).data0.n(2 * n(2) + 1)
  tn(4) = line3_value(no%).data(0).data0.n(2 * m(2))
  tn(5) = line3_value(no%).data(0).data0.n(2 * m(2) + 1)
  tl(0) = line3_value(t1%).data(0).data0.line_no(n(1))
  tl(1) = line3_value(t1%).data(0).data0.line_no(n(2))
  tl(2) = line3_value(no%).data(0).data0.line_no(m(2))
  s(0) = line3_value(no%).data(0).data0.para(m(0))
  s(1) = line3_value(no%).data(0).data0.para(m(1))
  s(2) = line3_value(no%).data(0).data0.para(m(2))
  v = line3_value(no%).data(0).data0.value
  If solve_multi_varity_equations(line3_value(t1%).data(0).data0.para(n(0)), _
   line3_value(t1%).data(0).data0.para(n(1)), line3_value(t1%).data(0).data0.para(n(2)), "0", _
    line3_value(t1%).data(0).data0.value, s(0), s(1), "0", s(2), v$, s(0), s(1), _
     s(2), v$) = False Then
      combine_three_three_line = set_equation(minus_string("x", v$, True, False), 0, temp_record)
       Exit Function
   End If
    If s(0) = "0" Then
    tp(0) = 0
     tp(1) = 0
    tn(0) = 0
     tn(1) = 0
    tl(0) = 0
   End If
   If s(1) = "0" Then
    tp(2) = 0
     tp(3) = 0
    tn(2) = 0
     tn(3) = 0
    tl(1) = 0
   End If
   If s(2) = "0" Then
    tp(4) = 0
     tp(5) = 0
    tn(4) = 0
     tn(5) = 0
    tl(2) = 0
   End If
  combine_three_three_line = set_three_line_value( _
   tp(0), tp(1), tp(2), tp(3), tp(4), tp(5), tn(0), tn(1), tn(2), tn(3), _
     tn(4), tn(5), tl(0), tl(1), tl(2), s(0), s(1), s(2), v, temp_record, _
      0, no_reduce, 0)
If combine_three_three_line > 0 Then
  Call set_level_(line3_value(no%).record_.no_reduce, 4)
If combine_three_three_line > 1 Then
   Exit Function
  End If
End If
 Next k%
For k% = 1 To last_tn12%
no% = tn12(k%)
temp_record.record_data = re
Call add_conditions_to_record(line3_value_, no%, 0, 0, temp_record.record_data.data0.condition_data)
  tp(0) = line3_value(t1%).data(0).data0.poi(2 * n(2))
  tp(1) = line3_value(t1%).data(0).data0.poi(2 * n(2) + 1)
  tp(2) = line3_value(t1%).data(0).data0.poi(2 * n(1))
  tp(3) = line3_value(t1%).data(0).data0.poi(2 * n(1) + 1)
  tp(4) = line3_value(no%).data(0).data0.poi(2 * m(2))
  tp(5) = line3_value(no%).data(0).data0.poi(2 * m(2) + 1)
  tn(0) = line3_value(t1%).data(0).data0.n(2 * n(2))
  tn(1) = line3_value(t1%).data(0).data0.n(2 * n(2) + 1)
  tn(2) = line3_value(t1%).data(0).data0.n(2 * n(1))
  tn(3) = line3_value(t1%).data(0).data0.n(2 * n(1) + 1)
  tn(4) = line3_value(no%).data(0).data0.n(2 * m(2))
  tn(5) = line3_value(no%).data(0).data0.n(2 * m(2) + 1)
  tl(0) = line3_value(t1%).data(0).data0.line_no(n(2))
  tl(1) = line3_value(t1%).data(0).data0.line_no(n(1))
  tl(2) = line3_value(no%).data(0).data0.line_no(m(2))
  s(0) = line3_value(no%).data(0).data0.para(m(0))
  s(1) = line3_value(no%).data(0).data0.para(m(1))
  s(2) = line3_value(no%).data(0).data0.para(m(2))
  v = line3_value(no%).data(0).data0.value
 If solve_multi_varity_equations(line3_value(t1%).data(0).data0.para(n(0)), _
   line3_value(t1%).data(0).data0.para(n(2)), line3_value(t1%).data(0).data0.para(n(1)), "0", _
    line3_value(t1%).data(0).data0.value, s(0), s(1), "0", s(2), v$, s(0), s(1), _
     s(2), v$) = False Then
      combine_three_three_line = set_equation(minus_string("x", v$, True, False), 0, temp_record)
       Exit Function
 End If
    If s(0) = "0" Then
    tp(0) = 0
     tp(1) = 0
    tn(0) = 0
     tn(1) = 0
    tl(0) = 0
   End If
   If s(1) = "0" Then
    tp(2) = 0
     tp(3) = 0
    tn(2) = 0
     tn(3) = 0
    tl(1) = 0
   End If
   If s(2) = "0" Then
    tp(4) = 0
     tp(5) = 0
    tn(4) = 0
     tn(5) = 0
    tl(2) = 0
   End If
  combine_three_three_line = set_three_line_value( _
   tp(0), tp(1), tp(2), tp(3), tp(4), tp(5), tn(0), tn(1), tn(2), tn(3), tn(4), _
     tn(5), tl(0), tl(1), tl(3), s(0), s(1), s(2), v, temp_record, 0, no_reduce, 0)
If combine_three_three_line > 0 Then
  Call set_level_(line3_value(no%).record_.no_reduce, 4)
If combine_three_three_line > 1 Then
   Exit Function
  End If
End If
 Next k%
For k% = 1 To last_tn21%
no% = tn21(k%)
temp_record.record_data = re
Call add_conditions_to_record(line3_value_, no%, 0, 0, temp_record.record_data.data0.condition_data)
  tp(0) = line3_value(t1%).data(0).data0.poi(2 * n(1))
  tp(1) = line3_value(t1%).data(0).data0.poi(2 * n(1) + 1)
  tp(2) = line3_value(t1%).data(0).data0.poi(2 * n(2))
  tp(3) = line3_value(t1%).data(0).data0.poi(2 * n(2) + 1)
  tp(4) = line3_value(no%).data(0).data0.poi(2 * m(1))
  tp(5) = line3_value(no%).data(0).data0.poi(2 * m(1) + 1)
  tn(0) = line3_value(t1%).data(0).data0.n(2 * n(1))
  tn(1) = line3_value(t1%).data(0).data0.n(2 * n(1) + 1)
  tn(2) = line3_value(t1%).data(0).data0.n(2 * n(2))
  tn(3) = line3_value(t1%).data(0).data0.n(2 * n(2) + 1)
  tn(4) = line3_value(no%).data(0).data0.n(2 * m(1))
  tn(5) = line3_value(no%).data(0).data0.n(2 * m(1) + 1)
  tl(0) = line3_value(t1%).data(0).data0.line_no(n(1))
  tl(1) = line3_value(t1%).data(0).data0.line_no(n(2))
  tl(2) = line3_value(no%).data(0).data0.line_no(m(1))
  s(0) = line3_value(no%).data(0).data0.para(m(0))
  s(1) = line3_value(no%).data(0).data0.para(m(2))
  s(2) = line3_value(no%).data(0).data0.para(m(1))
  v = line3_value(no%).data(0).data0.value
  If solve_multi_varity_equations(line3_value(t1%).data(0).data0.para(n(0)), _
   line3_value(t1%).data(0).data0.para(n(1)), line3_value(t1%).data(0).data0.para(n(2)), "0", _
    line3_value(t1%).data(0).data0.value, s(0), s(1), "0", s(2), v$, s(0), s(1), _
     s(2), v$) = False Then
        combine_three_three_line = set_equation(minus_string("x", v$, True, False), 0, temp_record)
       Exit Function
  End If
   If s(0) = "0" Then
    tp(0) = 0
     tp(1) = 0
    tn(0) = 0
     tn(1) = 0
    tl(0) = 0
   End If
   If s(1) = "0" Then
    tp(2) = 0
     tp(3) = 0
    tn(2) = 0
     tn(3) = 0
    tl(1) = 0
   End If
   If s(2) = "0" Then
    tp(4) = 0
     tp(5) = 0
    tn(4) = 0
     tn(5) = 0
    tl(2) = 0
   End If
  combine_three_three_line = set_three_line_value( _
   tp(0), tp(1), tp(2), tp(3), tp(4), tp(5), tn(0), tn(1), tn(2), _
     tn(3), tn(4), tn(5), tl(0), tl(1), tl(2), _
    s(0), s(1), s(2), v, temp_record, 0, no_reduce, 0)
If combine_three_three_line > 0 Then
  Call set_level_(line3_value(no%).record_.no_reduce, 4)
If combine_three_three_line > 1 Then
   Exit Function
  End If
End If
 Next k%
For k% = 1 To last_tn22%
no% = tn22(k%)
temp_record.record_data = re
Call add_conditions_to_record(line3_value_, no%, 0, 0, temp_record.record_data.data0.condition_data)
  tp(0) = line3_value(t1%).data(0).data0.poi(2 * n(2))
  tp(1) = line3_value(t1%).data(0).data0.poi(2 * n(2) + 1)
  tp(2) = line3_value(t1%).data(0).data0.poi(2 * n(1))
  tp(3) = line3_value(t1%).data(0).data0.poi(2 * n(1) + 1)
  tp(4) = line3_value(no%).data(0).data0.poi(2 * m(1))
  tp(5) = line3_value(no%).data(0).data0.poi(2 * m(1) + 1)
  tn(0) = line3_value(t1%).data(0).data0.n(2 * n(2))
  tn(1) = line3_value(t1%).data(0).data0.n(2 * n(2) + 1)
  tn(2) = line3_value(t1%).data(0).data0.n(2 * n(1))
  tn(3) = line3_value(t1%).data(0).data0.n(2 * n(1) + 1)
  tn(4) = line3_value(no%).data(0).data0.n(2 * m(1))
  tn(5) = line3_value(no%).data(0).data0.n(2 * m(1) + 1)
  tl(0) = line3_value(t1%).data(0).data0.line_no(n(2))
  tl(1) = line3_value(t1%).data(0).data0.line_no(n(1))
  tl(2) = line3_value(no%).data(0).data0.line_no(m(1))
  s(0) = line3_value(no%).data(0).data0.para(m(0))
  s(1) = line3_value(no%).data(0).data0.para(m(2))
  s(2) = line3_value(no%).data(0).data0.para(m(1))
  v = line3_value(no%).data(0).data0.value
  If solve_multi_varity_equations(line3_value(t1%).data(0).data0.para(n(0)), _
   line3_value(t1%).data(0).data0.para(n(2)), line3_value(t1%).data(0).data0.para(n(1)), "0", _
    line3_value(t1%).data(0).data0.value, s(0), s(1), "0", s(2), v$, s(0), s(1), _
     s(2), v$) = False Then
       combine_three_three_line = set_equation(minus_string("x", v$, True, False), 0, temp_record)
       Exit Function
   End If
   If s(0) = "0" Then
    tp(0) = 0
     tp(1) = 0
    tn(0) = 0
     tn(1) = 0
    tl(0) = 0
   End If
   If s(1) = "0" Then
    tp(2) = 0
     tp(3) = 0
    tn(2) = 0
     tn(3) = 0
    tl(1) = 0
   End If
   If s(2) = "0" Then
    tp(4) = 0
     tp(5) = 0
    tn(4) = 0
     tn(5) = 0
    tl(2) = 0
   End If
  combine_three_three_line = set_three_line_value( _
   tp(0), tp(1), tp(2), tp(3), tp(4), tp(5), tn(0), tn(1), tn(2), _
     tn(3), tn(4), tn(5), tl(0), tl(1), tl(2), _
    s(0), s(1), s(2), v, temp_record, 0, no_reduce, 0)
If combine_three_three_line > 0 Then
  Call set_level_(line3_value(no%).record_.no_reduce, 4)
If combine_three_three_line > 1 Then
   Exit Function
  End If
End If
 Next k%
combine_three_three_line_next2:
Next j%
combine_three_three_line_next1:
Next i%
End Function
Public Function combine_three_line_with_relation(ByVal t%, _
        ByVal start%, ByVal no_reduce As Byte) As Byte
Dim i%, j%, k%, no%, last_tn%
Dim n(2) As Integer
Dim m(2) As Integer
Dim n_(1) As Integer
Dim tn() As Integer
Dim v(1) As String
Dim re As relation_data0_type
If line3_value(t%).record_.no_reduce > 4 Then
 Exit Function
End If
 For i% = 0 To 2
  n(0) = i%
   n(1) = (i% + 1) Mod 3
    n(2) = (i% + 2) Mod 3
 For j% = 0 To 2
  m(0) = j%
   m(1) = (j% + 1) Mod 3
    m(2) = (j% + 2) Mod 3
  re.poi(2 * m(0)) = line3_value(t%).data(0).data0.poi(2 * n(0))
  re.poi(2 * m(0) + 1) = line3_value(t%).data(0).data0.poi(2 * n(0) + 1)
  re.poi(2 * m(1)) = -1
  Call search_for_relation(re, m(0), n_(0), 1)
  re.poi(2 * m(1)) = 30000
  Call search_for_relation(re, m(0), n_(1), 1)  '5.7
 last_tn% = 0
 For k% = n_(0) + 1 To n_(1)
 no% = Drelation(k%).data(0).record.data1.index.i(m(0))
 If no% > start% And Drelation(no%).record_.no_reduce < 4 Then
  'If is_two_record_related(relation_, no%, Drelation(no%).data(0).record, _
        line3_value_, t%, line3_value(t%).data(0).record) = False Then
  last_tn% = last_tn% + 1
ReDim Preserve tn(last_tn%) As Integer
tn(last_tn%) = no%
 End If
 'End If
 Next k%
 For k% = 1 To last_tn%
 no% = tn(k%)
  combine_three_line_with_relation = combine_relation_with_three_line_( _
     relation_, no%, line3_value_, t%, m(0), n(0))
If combine_three_line_with_relation > 1 Then
 Exit Function
End If
Next k%
Next j%
Next i%
End Function
Public Function combine_two_angle0(ByVal A1%, ByVal A2%) As Integer
Dim i%, A%
'合并
If angle(A1%).data(0).poi(1) = angle(A2%).data(0).poi(1) Then
 For i% = 0 To 1
  If angle(A1%).data(0).line_no(i%) = angle(A2%).data(0).line_no((i% + 1) Mod 2) Then
   A% = angle_number(m_lin(angle(A1%).data(0).line_no((i% + 1) Mod 2)).data(0).data0.poi(angle(A1%).data(0).te((i% + 1) Mod 2)), _
             angle(A1%).data(0).poi(1), _
                m_lin(angle(A2%).data(0).line_no(i%)).data(0).data0.poi(angle(A2%).data(0).te(i%)), 0, 0)
     If A% <> 0 Then
      combine_two_angle0 = Abs(A%)
     End If
       Exit Function
  End If
 Next i%
End If
End Function
Public Function combine_two_line_with_three_line(ByVal t_l%, _
       ByVal start%, ByVal no_reduce As Byte) As Byte
Dim k%, l%, i%, no%
Dim n(1) As Integer
Dim m(2) As Integer
Dim n_(1) As Integer
Dim th_l As line3_value_data0_type
Dim tn() As Integer
Dim last_tn As Integer
If two_line_value(t_l%).record_.no_reduce > 4 Then
 Exit Function
End If
For k% = 0 To 1
n(0) = k%
n(1) = (k% + 1) Mod 2
For l% = 0 To 2
m(0) = l%
m(1) = (l% + 1) Mod 3
m(2) = (l% + 2) Mod 3
th_l.poi(2 * m(0)) = two_line_value(t_l%).data(0).data0.poi(2 * n(0))
th_l.poi(2 * m(0) + 1) = two_line_value(t_l%).data(0).data0.poi(2 * n(0) + 1)
th_l.poi(2 * m(1)) = -1
Call search_for_line3_value(th_l, l%, n_(0), 1)
th_l.poi(2 * m(1)) = 30000
Call search_for_line3_value(th_l, l%, n_(1), 1)  '5.7
last_tn = 0
For i% = n_(0) + 1 To n_(1)
no% = line3_value(i%).data(0).record.data1.index.i(l%)
If line3_value(no%).record_.no_reduce < 4 And _
 no% > start% Then
' If is_two_record_related(line3_value_, no%, line3_value(no%).data(0).record, _
        two_line_value_, t_l%, two_line_value(t_l%).data(0).record) = False Then
last_tn = last_tn + 1
ReDim Preserve tn(last_tn) As Integer
tn(last_tn) = no%
End If
'End If
Next i%
For i% = 1 To last_tn
combine_two_line_with_three_line = _
 combine_three_three_line_(two_line_value_, t_l%, line3_value_, no%, _
     n(0), m(0), 0)
 If combine_two_line_with_three_line > 1 Then
   Exit Function
  End If
Next i%
Next l% '
Next k%
End Function
Public Function combine_two_angle3(a3_v1 As angle3_value_data0_type, _
                     A3_v2 As angle3_value_data0_type, k%, l%, _
                       out_A3_v As angle3_value_data0_type) As Boolean
Dim A1(2) As Integer
Dim A2(2) As Integer
Dim A(3) As Integer
Dim para(3) As Integer
Dim para1(2) As String
Dim para2(2) As String
Dim v(1) As String
Dim n(2) As Integer
Dim m(2) As Integer
Dim i%, j%
For i% = 0 To 2
A1(i%) = a3_v1.angle(i%)
para1(i%) = a3_v1.para(i%)
A2(i%) = A3_v2.angle(i%)
para2(i%) = A3_v2.para(i%)
Next i%
v(0) = a3_v1.value
v(1) = A3_v2.value
If k% > 0 And l% > 0 Then
   n(0) = k%
    n(1) = (k% + 1) Mod 3
     n(2) = (k% + 2) Mod 3
   m(0) = l%
    m(1) = (l% + 1) Mod 3
     m(2) = (l% + 2) Mod 3
Else
For i% = 0 To 2
 For j% = 0 To 2
  If A1(i%) = A2(j%) Then
   n(0) = i%
    n(1) = (i% + 1) Mod 3
     n(2) = (i% + 2) Mod 3
   m(0) = j%
    m(1) = (j% + 1) Mod 3
     m(2) = (j% + 2) Mod 3
      GoTo combine_two_angle3_mark0
  End If
 Next j%
Next i%
Exit Function
End If
combine_two_angle3_mark0:
A(0) = A1(n(1))
A(1) = A1(n(2))
A(2) = A2(m(1))
A(3) = A2(m(2))
para(0) = time_string(para1(n(1)), para2(m(0)), True, False)
para(1) = time_string(para1(n(2)), para2(m(0)), True, False)
para(2) = time_string(para2(m(1)), para1(n(0)), True, False)
para(2) = time_string("-1", para(2), True, False)
para(3) = time_string(para2(m(2)), para1(n(0)), True, False)
para(3) = time_string("-1", para(3), True, False)
v(0) = minus_string( _
     time_string(v(0), para2(m(0)), False, False), _
      time_string(v(1), para1(n(0)), False, False), True, False)
If A(0) = A(2) Then
 para(0) = add_string(para(0), para(2), True, False)
  para(2) = "0"
ElseIf A(0) = A(3) Then
 para(0) = add_string(para(0), para(3), True, False)
  para(3) = "0"
End If
If A(1) = A(2) Then
 para(1) = add_string(para(1), para(2), True, False)
  para(2) = "0"
ElseIf A(1) = A(3) Then
 para(1) = add_string(para(1), para(3), True, False)
  para(3) = "0"
End If
For i% = 0 To 2
 If para(i%) = "0" Then
  For j% = i% To 2
   para(j%) = para(j% + 1)
    A(j%) = A(j% + 1)
  Next j%
   para(3) = "0"
   A(3) = 0
 End If
Next i%
If para(3) = "0" Then
out_A3_v.angle(0) = A(0)
out_A3_v.angle(1) = A(1)
out_A3_v.angle(2) = A(2)
out_A3_v.para(0) = para(0)
out_A3_v.para(0) = para(1)
out_A3_v.para(0) = para(2)
out_A3_v.value = v(0)
combine_two_angle3 = True
End If
End Function
Public Function combine_two_line3(line3_v1 As line3_value_data0_type, _
          line3_v2 As line3_value_data0_type, k%, l%, _
           out_line3_v As line3_value_data0_type) As Boolean
Dim p1(5) As Integer
Dim tn1(5) As Integer
Dim tl1(2) As Integer
Dim p2(5) As Integer
Dim tn2(5) As Integer
Dim tl2(2) As Integer
Dim p(7) As Integer
Dim tn(7) As Integer
Dim tl(3) As Integer
Dim para(3) As Integer
Dim para1(2) As String
Dim para2(2) As String
Dim v(1) As String
Dim n(2) As Integer
Dim m(2) As Integer
Dim i%, j%
For i% = 0 To 2
p1(2 * i%) = line3_v1.poi(2 * i%)
p1(2 * i% + 1) = line3_v1.poi(2 * i% + 1)
tn1(2 * i%) = line3_v1.n(2 * i%)
tn1(2 * i% + 1) = line3_v1.n(2 * i% + 1)
p2(2 * i%) = line3_v2.poi(2 * i%)
p2(2 * i% + 1) = line3_v2.poi(2 * i% + 1)
tn2(2 * i%) = line3_v2.n(2 * i%)
tn2(2 * i% + 1) = line3_v2.n(2 * i% + 1)
tl1(i%) = line3_v1.line_no(i%)
para1(i%) = line3_v1.para(i%)
tl2(i%) = line3_v2.line_no(i%)
para2(i%) = line3_v2.para(i%)
Next i%
v(0) = line3_v1.value
v(1) = line3_v2.value
If k% > 0 And l% > 0 Then
   n(0) = k%
    n(1) = (k% + 1) Mod 3
     n(2) = (k% + 2) Mod 3
   m(0) = l%
    m(1) = (l% + 1) Mod 3
     m(2) = (l% + 2) Mod 3
Else
For i% = 0 To 2
 For j% = 0 To 2
  If p1(2 * i%) = p2(2 * j%) And _
    p1(2 * i% + 1) = p2(2 * j% + 1) Then
   n(0) = i%
    n(1) = (i% + 1) Mod 3
     n(2) = (i% + 2) Mod 3
   m(0) = j%
    m(1) = (j% + 1) Mod 3
     m(2) = (j% + 2) Mod 3
      GoTo combine_two_line3_mark0
  End If
 Next j%
Next i%
Exit Function
End If
combine_two_line3_mark0:
p(0) = p1(2 * n(1))
p(1) = p1(2 * n(1) + 1)
p(2) = p1(2 * n(2))
p(3) = p1(2 * n(2) + 1)
p(4) = p2(2 * m(1))
p(5) = p2(2 * m(1) + 1)
p(6) = p2(2 * m(2))
p(7) = p2(2 * m(2) + 1)
tl(0) = tl1(n(1))
tl(1) = tl1(n(2))
tl(2) = tl2(m(1))
tl(3) = tl2(m(2))
tn(0) = tn1(2 * n(1))
tn(1) = tn1(2 * n(1) + 1)
tn(2) = tn1(2 * n(2))
tn(3) = tn1(2 * n(2) + 1)
tn(4) = tn2(2 * m(1))
tn(5) = tn2(2 * m(1) + 1)
tn(6) = tn2(2 * m(2))
tn(7) = tn2(2 * m(2) + 1)
para(0) = time_string(para1(n(1)), para2(m(0)), True, False)
para(1) = time_string(para1(n(2)), para2(m(0)), True, False)
para(2) = time_string(para2(m(1)), para1(n(0)), True, False)
para(2) = time_string("-1", para(2), True, False)
para(3) = time_string(para2(m(2)), para1(n(0)), True, False)
para(3) = time_string("-1", para(3), True, False)
v(0) = minus_string( _
     time_string(v(0), para2(m(0)), False, False), _
      time_string(v(1), para1(n(0)), False, False), True, False)
If p(0) = p(4) And p(1) = p(5) Then
 para(0) = add_string(para(0), para(2), True, False)
  para(2) = "0"
ElseIf p(0) = p(6) And p(1) = p(7) Then
 para(0) = add_string(para(0), para(3), True, False)
  para(3) = "0"
End If
If p(2) = p(4) And p(3) = p(5) Then
 para(1) = add_string(para(1), para(2), True, False)
  para(2) = "0"
ElseIf p(2) = p(6) And p(3) And p(7) Then
 para(1) = add_string(para(1), para(3), True, False)
  para(3) = "0"
End If
For i% = 0 To 2
 If para(i%) = "0" Then
  For j% = i% To 2
   para(j%) = para(j% + 1)
    p(2 * j%) = p(2 * (j% + 1))
     p(2 * j% + 1) = p(2 * j% + 3)
    tn(2 * j%) = tn(2 * (j% + 1))
     tn(2 * j% + 1) = tn(2 * j% + 3)
    tl(j%) = tl(j% + 1)
  Next j%
   para(3) = "0"
   p(6) = 0
    p(7) = 0
   tn(6) = 0
    tn(7) = 0
   tl(3) = 0
 End If
Next i%
If para(3) = "0" Then
For i% = 0 To 2
out_line3_v.poi(2 * i%) = p(2 * i%)
out_line3_v.poi(2 * i% + 1) = p(2 * i% + 1)
out_line3_v.n(2 * i%) = tn(2 * i%)
out_line3_v.n(2 * i% + 1) = tn(2 * i% + 1)
out_line3_v.line_no(i%) = tl(i%)
Next i%
out_line3_v.value = v(0)
combine_two_line3 = True
End If
End Function
Public Function combine_tri_function_with_item0(ByVal n%, no_reduce As Byte) As Byte
Dim it As item0_data_type
Dim i%, j%, k%, no%
Dim tn() As Integer
Dim n_(1) As Integer
Dim m(1) As Integer
Dim temp_record0 As record_type0
For i% = 0 To 1
m(0) = i%
m(1) = (i% + 1) Mod 2
it.poi(2 * m(0)) = tri_function(n%).data(0).A
it.poi(2 * m(0) + 1) = -6
Call search_for_item0(it, i%, n_(0), 1)
it.poi(2 * m(0) + 1) = 0
Call search_for_item0(it, i%, n_(1), 1)
k% = 0
For j% = n_(0) + 1 To n_(1)
no% = item0(j%).data(0).index(i%)
ReDim Preserve tn(k%) As Integer
tn(k%) = no%
k% = k% + 1
Next j%
For j% = 0 To k% - 1
no% = tn(j%)
temp_record0.condition_data.condition_no = 1
temp_record0.condition_data.condition(1).ty = tri_function_
temp_record0.condition_data.condition(1).no = n%
combine_tri_function_with_item0 = combine_tri_function_with_item0_( _
      tri_function(n%).data(0).A, no%, i%, temp_record0)
    If combine_tri_function_with_item0 > 1 Then
     Exit Function
    End If
Next j%
Next i%

End Function

Public Function combine_item0_value_with_item0(it%, no_reduce As Byte) As Byte
Dim i%, tn%
Dim temp_record0 As record_type0
temp_record0.condition_data = item0(it%).data(0).record_for_value.data0.condition_data
For i% = 1 To last_conditions.last_cond(1).item0_no
 If item0(i%).data(0).sig = "i" Then
  If item0(i%).data(0).poi(0) = it% Then
     tn% = 0
     combine_item0_value_with_item0 = set_item0(item0(it%).data(0).poi(1), _
         item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), 0, "i", _
           0, 0, 0, 0, 0, 0, "1", "1", item0(it%).data(0).value, "", _
             "1", 0, record_data0.data0.condition_data, 0, tn%, no_reduce, _
               0, condition_data0, False)
      If combine_item0_value_with_item0 > 1 Then
       Exit Function
      End If
'     If tn% > 0 Then
'      combine_item0_value_with_item0 = combine_item_with_general_string_( _
        it%, tn%)
'      If combine_item0_value_with_item0 > 1 Then
'       Exit Function
'      End If
'     End If
  ElseIf item0(i%).data(0).poi(1) = it% Then
     combine_item0_value_with_item0 = set_item0(item0(it%).data(0).poi(0), _
         item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), 0, "i", _
           0, 0, 0, 0, 0, 0, "1", "1", item0(it%).data(0).value, "", "1", _
            0, record_data0.data0.condition_data, 0, tn%, no_reduce, 0, condition_data0, False)
      If combine_item0_value_with_item0 > 1 Then
       Exit Function
      End If
 '    If tn% > 0 Then
 '     combine_item0_value_with_item0 = combine_item_with_general_string_( _
        it%, tn%)
 '     If combine_item0_value_with_item0 > 1 Then
 '      Exit Function
 '     End If
 '    End If
  ElseIf item0(i%).data(0).poi(2) = it% Then
       tn% = 0
       combine_item0_value_with_item0 = set_item0(item0(it%).data(0).poi(0), _
         item0(it%).data(0).poi(1), item0(it%).data(0).poi(3), 0, "i", _
           0, 0, 0, 0, 0, 0, "1", "1", item0(it%).data(0).value, "", "1", _
            0, record_data0.data0.condition_data, 0, tn%, no_reduce, 0, condition_data0, False)
      If combine_item0_value_with_item0 > 1 Then
       Exit Function
      End If
  '   If tn% > 0 Then
  '    combine_item0_value_with_item0 = combine_item_with_general_string_( _
        it%, tn%)
  '    If combine_item0_value_with_item0 > 1 Then
  '     Exit Function
  '    End If
  '   End If
  ElseIf item0(i%).data(0).poi(3) = it% Then
     tn% = 0
     combine_item0_value_with_item0 = set_item0(item0(it%).data(0).poi(0), _
         item0(it%).data(0).poi(1), item0(it%).data(0).poi(2), 0, "i", _
           0, 0, 0, 0, 0, 0, "1", "1", item0(it%).data(0).value, "", "1", _
            0, record_data0.data0.condition_data, 0, tn%, no_reduce, 0, _
              condition_data0, False)
      If combine_item0_value_with_item0 > 1 Then
       Exit Function
      End If
  '   If tn% > 0 Then
  '    combine_item0_value_with_item0 = combine_item_with_general_string_( _
  '      it%, tn%)
  '    If combine_item0_value_with_item0 > 1 Then
  '     Exit Function
  '    End If
  '   End If
  End If
 End If
Next i%
End Function

Public Function combine_eangle_with_item0(ByVal tA%, no_reduce As Byte) As Byte
Dim i%, j%, k%
Dim tn() As Integer
Dim last_tn%
Dim n_(1) As Integer
Dim m(1) As Integer
Dim n(1) As Integer
Dim ite As item0_data_type
Dim temp_record As total_record_type
If angle3_value(tA%).data(0).data0.type = eangle_ Then
temp_record.record_data.data0.condition_data.condition_no = 1
temp_record.record_data.data0.condition_data.condition(1).ty = angle3_value_
temp_record.record_data.data0.condition_data.condition(1).no = tA%
For i% = 0 To 1
m(0) = i%
m(1) = (i% + 1) Mod 2
For j% = 0 To 1
n(0) = j%
n(1) = (j% + 1) Mod 2
ite.poi(2 * m(0)) = angle3_value(tA%).data(0).data0.angle(n(0))
ite.poi(2 * m(0) + 1) = -5
Call search_for_item0(ite, m(0), n_(0), 0)
ite.poi(2 * m(0) + 1) = 0
Call search_for_item0(ite, m(0), n_(1), 0)
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
  last_tn% = last_tn% + 1
   ReDim Preserve tn(last_tn%) As Integer
   tn(last_tn%) = item0(k%).data(0).index(m(0))
Next k%
For k% = 1 To last_tn%
      combine_eangle_with_item0 = combine_eangle_with_item0_(tA%, _
       tn(k%), n(0), m(0), temp_record.record_data, no_reduce)
      If combine_eangle_with_item0 > 1 Then
       Exit Function
      End If
Next k%
Next j%
Next i%
End If
End Function
Private Function combine_relation_with_line_(p() As Integer, v() As String, _
         l_p() As Integer, out_v As String, ty As Byte) As Boolean
If p(0) = l_p(0) And p(1) = l_p(1) Then
 If ty = 0 Then
  out_v = "1"
 ElseIf ty = 1 Then
  out_v = v(0)
 ElseIf ty = 2 Then
  out_v = v(1)
 End If
ElseIf p(2) = l_p(0) And p(3) = l_p(1) Then
 If ty = 0 Then
 If v(0) <> "" Then
   If v(0) <> "0" Then
    out_v = divide_string("1", v(0), True, False)
   Else
    out_v = ""
   End If
 Else
  out_v = ""
 End If
 ElseIf ty = 1 Then
  out_v = "1"
 ElseIf ty = 2 Then
  If v(0) <> "" And v(1) <> "" Then
   out_v = divide_string(v(1), v(0), True, False)
  Else
   out_v = ""
  End If
 End If
ElseIf p(4) = l_p(0) And p(5) = l_p(1) Then
 If ty = 0 Then
 If v(1) <> "" Then
   If v(1) <> "0" Then
    out_v = divide_string("1", v(1), True, False)
   Else
    out_v = ""
   End If
 Else
   out_v = ""
 End If
 ElseIf ty = 1 Then
  If v(0) <> "" And v(1) <> "" Then
    out_v = divide_string(v(0), v(1), True, False)
  Else
    out_v = ""
  End If
 ElseIf ty = 2 Then
   out_v = "1"
 End If
Else
 out_v = ""
End If
 If out_v <> "" Then
 combine_relation_with_line_ = True
 End If
End Function
Public Function combine_relation_with_item0_(p() As Integer, n() As Integer, _
              l() As Integer, v() As String, in_item As item0_data_type, _
               out_item As item0_data_type, para, no_reduce As Byte) As Boolean
Dim tp(1) As Integer
Dim ite As item0_data_type
Dim t_para1 As String
Dim t_para2 As String
Dim ty(1) As Boolean
  tp(0) = in_item.poi(0)
  tp(1) = in_item.poi(1)
If tp(0) > 0 And tp(1) > 0 Then
 ty(0) = combine_relation_with_line_(p(), v(), tp(), t_para1, 0)
End If
  tp(0) = in_item.poi(2)
  tp(1) = in_item.poi(3)
If tp(0) > 0 And tp(1) > 0 Then
  ty(1) = combine_relation_with_line_(p(), v(), tp(), t_para2, 0)
End If
If ty(0) And ty(1) Then
    If in_item.sig = "*" Then
    out_item = set_item0_(p(0), p(1), p(0), p(1), "*", n(0), n(1), n(0), n(1), _
      l(0), l(0), "", "", "", 0, "", condition_data0)
       para = time_string(t_para1, t_para2, True, False)
    ElseIf in_item.sig = "/" Then
       out_item = ite
        para = divide_string(t_para1, t_para2, True, False)
    End If
    combine_relation_with_item0_ = True
ElseIf ty(0) Then
    out_item = set_item0_(p(0), p(1), in_item.poi(2), in_item.poi(3), in_item.sig, _
                    n(0), n(1), in_item.n(2), in_item.n(3), l(0), in_item.line_no(1), _
                      "", "", "", 0, "", condition_data0)
    para = t_para1
    combine_relation_with_item0_ = True
ElseIf ty(1) Then
    out_item = set_item0_(in_item.poi(0), in_item.poi(1), p(0), p(1), in_item.sig, _
                    in_item.n(0), in_item.n(1), n(0), n(1), in_item.line_no(0), l(0), _
                      "", "", "", 0, "", condition_data0)
       If in_item.sig = "*" Or in_item.sig = "~" Then
        para = t_para2
       ElseIf in_item.sig = "/" Then
        para = divide_string("1", t_para2, True, False)
       End If
    combine_relation_with_item0_ = True
End If
End Function

Public Function combine_angle_value_with_item0(ByVal A%, ByVal value$, re As record_data_type) As Byte
Dim it As item0_data_type
Dim i%, j%, k%, no%
Dim tn() As Integer
Dim n_(1) As Integer
Dim m(1) As Integer
Dim temp_record0 As record_type0
For i% = 0 To 1
m(0) = i%
m(1) = (i% + 1) Mod 2
it.poi(2 * m(0)) = A%
it.poi(2 * m(0) + 1) = 0
Call search_for_item0(it, i%, n_(0), 1)
it.poi(2 * m(0) + 1) = -6
Call search_for_item0(it, i%, n_(1), 1)
k% = 0
For j% = n_(0) + 1 To n_(1)
no% = item0(j%).data(0).index(i%)
ReDim Preserve tn(k%) As Integer
tn(k%) = no%
k% = k% + 1
Next j%
For j% = 0 To k% - 1
no% = tn(j%)
temp_record0.condition_data.condition_no = 1
temp_record0.condition_data.condition(1).ty = angle3_value_
temp_record0.condition_data.condition(1).no = angle(A%).data(0).value_no
combine_angle_value_with_item0 = combine_tri_function_with_item0_( _
       A%, no%, i%, temp_record0)
    If combine_angle_value_with_item0 > 1 Then
     Exit Function
    End If
Next j%
Next i%

End Function

Public Function combine_tri_function_with_item0_(ByVal A%, ByVal it%, k%, re As record_type0) As Byte
Dim n(1) As Integer
Dim temp_record0 As record_type0
Dim ts$
Dim T_V$
Dim tri_n%
n(0) = k%
n(1) = (k + 1) Mod 2
If angle(A%).data(0).value <> "" Then
   ts$ = angle(A%).data(0).value
     If item0(it%).data(0).poi(2 * n(0) + 1) = -1 Then
      T_V$ = sin_(ts$, 0)
     ElseIf item0(it%).data(0).poi(2 * n(0) + 1) = -2 Then
      T_V$ = cos_(ts$, 0)
     ElseIf item0(it%).data(0).poi(2 * n(0) + 1) = -3 Then
      T_V$ = tan_(ts$, 0)
     ElseIf item0(it%).data(0).poi(2 * n(0) + 1) = -4 Then
      T_V$ = divide_string("1", tan_(ts$, 0), True, False)
     Else
      Exit Function
     End If
     temp_record0.condition_data.condition_no = 1
     temp_record0.condition_data.condition(1).ty = angle3_value_
     temp_record0.condition_data.condition(1).no = angle(A%).data(0).value_no
ElseIf re.condition_data.condition(1).ty = tri_function_ Then
    tri_n% = re.condition_data.condition(1).no
     If item0(it%).data(0).poi(2 * n(0) + 1) = -1 Then
      T_V$ = tri_function(tri_n%).data(0).sin_value
     ElseIf item0(it%).data(0).poi(2 * n(0) + 1) = -2 Then
      T_V$ = tri_function(tri_n%).data(0).cos_value
     ElseIf item0(it%).data(0).poi(2 * n(0) + 1) = -3 Then
      T_V$ = tri_function(tri_n%).data(0).tan_value
     ElseIf item0(it%).data(0).poi(2 * n(0) + 1) = -4 Then
      T_V$ = tri_function(tri_n%).data(0).ctan_value
     Else
      Exit Function
     End If
     temp_record0 = re
Else
 Exit Function
End If
If T_V$ = "F" Then
    Exit Function
End If
If item0(it%).data(0).poi(2 * n(1)) = 0 Then
  If item0(it%).data(0).value <> "" Then
   Exit Function
  Else
      item0(it%).data(0).value = T_V$
      item0(it%).data(0).record_for_value.data0.condition_data = re.condition_data
     combine_tri_function_with_item0_ = combine_item_with_general_string_( _
        it%, -1)
     If combine_tri_function_with_item0_ > 1 Then
      Exit Function
     End If
  End If
Else 'v=""
 If item0(it%).data(0).sig = "*" Then
  '   temp_record0.condition_data = re.condition_data
      temp_record0.para(0) = T_V$
       temp_record0.para(1) = "1"
    combine_tri_function_with_item0_ = set_item0(item0(it%).data(0).poi(2 * n(1)), _
     item0(it%).data(0).poi(2 * n(1) + 1), 0, 0, "~", _
      item0(it%).data(0).n(2 * n(1)), item0(it%).data(0).n(2 * n(1) + 1), _
       item0(it%).data(0).line_no(n(1)), 0, 0, 0, "1", "1", T_V$, _
        "", "1", 0, temp_record0.condition_data, it%, 0, 0, 0, condition_data0, False)
     If combine_tri_function_with_item0_ > 1 Then
      Exit Function
     End If
 ElseIf item0(it%).data(0).sig = "/" Then
  If k% = 0 Then
    ' temp_record0.condition_data = re.condition_data
      temp_record0.para(0) = T_V$
       temp_record0.para(1) = "1"
    combine_tri_function_with_item0_ = set_item0(0, 0, item0(it%).data(0).poi(2 * n(1)), _
     item0(it%).data(0).poi(2 * n(1) + 1), "/", 0, 0, 0, _
      item0(it%).data(0).n(2 * n(1)), item0(it%).data(0).n(2 * n(1) + 1), _
       item0(it%).data(0).line_no(n(1)), "1", "1", T_V$, "", _
        "1", 0, temp_record0.condition_data, it%, 0, 0, 0, condition_data0, False)
     If combine_tri_function_with_item0_ > 1 Then
      Exit Function
     End If
  Else
      temp_record0.para(0) = divide_string("1", T_V$, True, False)
       temp_record0.para(1) = "1"
        'temp_record0.condition_data = re.condition_data
    combine_tri_function_with_item0_ = set_item0(item0(it%).data(0).poi(2 * n(1)), _
     item0(it%).data(0).poi(2 * n(1) + 1), 0, 0, "~", _
      item0(it%).data(0).n(2 * n(1)), item0(it%).data(0).n(2 * n(1) + 1), _
       item0(it%).data(0).line_no(n(1)), 0, 0, 0, "1", "1", temp_record0.para(0), "", _
        "1", 0, temp_record0.condition_data, it%, 0, 0, 0, condition_data0, False)
     If combine_tri_function_with_item0_ > 1 Then
      Exit Function
     End If
  End If
 End If
End If
End Function

Public Function combine_eangle_with_item0_(eA%, it%, k%, j%, re As record_data_type, no_reduce As Byte) As Byte
Dim n(1) As Integer
Dim temp_record0 As record_type0
Dim temp_record As total_record_type
Dim tv$
n(0) = k%
n(1) = (k% + 1) Mod 2
temp_record0.condition_data = re.data0.condition_data
 If angle(angle3_value(eA%).data(0).data0.angle(n(1))).data(0).value <> "" Then
  If item0(it%).data(0).poi(2 * j% + 1) = -1 Then
   tv$ = sin_(angle(angle3_value(eA%).data(0).data0.angle(n(1))).data(0).value, 0)
  ElseIf item0(it%).data(0).poi(2 * j% + 1) = -2 Then
   tv$ = cos_(angle(angle3_value(eA%).data(0).data0.angle(n(1))).data(0).value, 0)
  ElseIf item0(it%).data(0).poi(2 * j% + 1) = -3 Then
   tv$ = tan_(angle(angle3_value(eA%).data(0).data0.angle(n(1))).data(0).value, 0)
  ElseIf item0(it%).data(0).poi(2 * j% + 1) = -4 Then
   tv$ = tan_(angle(angle3_value(eA%).data(0).data0.angle(n(1))).data(0).value, 0)
    tv$ = divide_string("1", tv$, True, False)
  End If
 End If
If tv$ <> "F" And tv$ <> "" Then
  temp_record.record_data = re
  If j% = 0 Then
    If item0(it%).data(0).poi(2) = 0 Then
       Call add_conditions_to_record(angle3_value_, _
             angle(angle3_value(eA%).data(0).data0.angle(n(1))).data(0).value_no, 0, 0, _
               temp_record.record_data.data0.condition_data)
       combine_eangle_with_item0_ = set_item0_value(it%, 0, 0, "0", "0", tv$, "", _
                  0, temp_record.record_data.data0.condition_data)
    Else
    End If
  Else
    If item0(it%).data(0).poi(0) = 0 Then
    Else
    End If
  End If
Else
If j% = 0 Then
combine_eangle_with_item0_ = set_item0(angle3_value(eA%).data(0).data0.angle(n(1)), _
 item0(it%).data(0).poi(1), item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), _
  item0(it%).data(0).sig, 0, 0, item0(it%).data(0).n(2), item0(it%).data(0).n(3), _
   0, item0(it%).data(0).line_no(1), "1", "1", "1", "", "1", 0, temp_record0.condition_data, _
    it%, 0, no_reduce, 0, condition_data0, False)
  If combine_eangle_with_item0_ > 1 Then
   Exit Function
  End If
Else
combine_eangle_with_item0_ = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
  angle3_value(eA%).data(0).data0.angle(n(1)), item0(it%).data(0).poi(3), item0(it%).data(0).sig, _
   item0(it%).data(0).n(0), item0(it%).data(0).n(1), 0, 0, _
   item0(it%).data(0).line_no(0), 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, _
     it%, 0, no_reduce, 0, condition_data0, False)
  If combine_eangle_with_item0_ > 1 Then
   Exit Function
  End If
End If
End If
End Function

Public Function combine_item_with_eangle(ByVal it%, no_reduce As Byte)
Dim i%, j%, k%, no%
Dim n_(1) As Integer
Dim n(1) As Integer
Dim m(1) As Integer
Dim tn() As Integer
Dim last_tn%
Dim tA As angle3_value_data0_type
Dim temp_record As total_record_type
For i% = 0 To 1
n(0) = i%
n(1) = (i% + 1) Mod 2
 If item0(it%).data(0).poi(2 * n(0) + 1) > -5 And _
      item0(it%).data(0).poi(2 * n(0) + 1) < 0 Then
 For j% = 0 To 1
  m(0) = j%
  m(1) = (j% + 1) Mod 2
  tA.angle(m(0)) = item0(it%).data(0).poi(2 * n(0))
  tA.angle(m(1)) = -1
  Call search_for_three_angle_value(tA, m(0), n_(0), 1)
  tA.angle(m(1)) = 30000
  Call search_for_three_angle_value(tA, m(0), n_(1), 1)
  'm(1) = (j% + 1) Mod 2
  last_tn% = 0
  For k% = n_(0) + 1 To n_(1)
   no% = angle3_value(k%).data(0).record.data1.index.i(m(0))
   If angle3_value(no%).data(0).data0.type = eangle_ Then
   last_tn% = last_tn% + 1
   ReDim Preserve tn(last_tn%) As Integer
   tn(last_tn%) = no%
   End If
  Next k%
  For k% = 1 To last_tn%
  no% = tn(k%)
  If angle3_value(no%).data(0).data0.type = eangle_ Then
  temp_record.record_data.data0.condition_data.condition_no = 1
  temp_record.record_data.data0.condition_data.condition(1).ty = angle3_value_
  temp_record.record_data.data0.condition_data.condition(1).no = no%
  combine_item_with_eangle = combine_eangle_with_item0_(no%, it%, m(0), n(0), _
   temp_record.record_data, no_reduce)
  If combine_item_with_eangle > 1 Then
   Exit Function
  End If
  End If
  Next k%
 Next j%
 End If
Next i%
End Function
Public Function combine_line_value_with_item_(ByVal lv%, ByVal it%, ByVal j%, no_reduce As Byte) As Byte
Dim m(1) As Integer
Dim it_(1) As Integer
Dim temp_record1 As condition_data_type
Dim temp_record As total_record_type
Dim temp_record_data As record_data_type
If item0(it%).data(0).no_reduce = False Then
temp_record_data.data0.condition_data.condition_no = 1
temp_record_data.data0.condition_data.condition(1).ty = line_value_
temp_record_data.data0.condition_data.condition(1).no = lv%
temp_record_data.data0.theorem_no = 1
If item0(it%).data(0).no_reduce = False Then
If item0(it%).data(0).sig = "*" Then
If line_value(lv%).data(0).data0.poi(0) = item0(it%).data(0).poi(2 * ((j% + 1) Mod 2)) And _
    line_value(lv%).data(0).data0.poi(1) = item0(it%).data(0).poi(2 * ((j% + 1) Mod 2) + 1) And j% < 2 Then
      If item0(it%).data(0).value = "" Then '平方项
        item0(it%).data(0).value = line_value(lv%).data(0).data0.squar_value
         item0(it%).data(0).record_for_value.data0.condition_data = _
                                    temp_record_data.data0.condition_data
            combine_line_value_with_item_ = combine_item_with_general_string_(it%, -1)
      End If
Else
  If item0(it%).data(0).value <> "" Then
   Call add_record_to_record(item0(it%).data(0).record_for_value.data0.condition_data, _
                                                      temp_record_data.data0.condition_data)
    temp_record.record_data = temp_record_data
    If j% = 0 Then
     combine_line_value_with_item_ = set_element_value(it%, 1, _
       divide_string(item0(it%).data(0).value, line_value(lv%).data(0).data0.value, True, False), _
         temp_record, no_reduce)
         If combine_line_value_with_item_ > 1 Then
          Exit Function
         End If
          item0(it%).data(0).no_reduce = True
           Call combine_item_with_general_string_(it%, -2)
    ElseIf j% = 1 Then
     combine_line_value_with_item_ = set_element_value(it%, 0, _
       divide_string(item0(it%).data(0).value, line_value(lv%).data(0).data0.value, True, False), _
         temp_record, no_reduce)
         If combine_line_value_with_item_ > 1 Then
          Exit Function
         End If
        item0(it%).data(0).no_reduce = True
           Call combine_item_with_general_string_(it%, -2)
    Else
     If item0(it%).data(0).diff_type = 3 Or item0(it%).data(0).diff_type = 5 Then
      combine_line_value_with_item_ = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
             item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), "*", _
              item0(it%).data(0).n(0), item0(it%).data(0).n(1), item0(it%).data(0).n(0), _
               item0(it%).data(0).n(1), item0(it%).data(0).line_no(0), item0(it%).data(0).line_no(0), _
                "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, m(0), _
                 no_reduce, 0, condition_data0, False)
             If combine_line_value_with_item_ > 1 Then
               Exit Function
             End If
      combine_line_value_with_item_ = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
             0, 0, "~", item0(it%).data(0).n(0), item0(it%).data(0).n(1), _
              0, 0, item0(it%).data(0).line_no(0), 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, _
                0, m(1), no_reduce, 0, condition_data0, False)
             If combine_line_value_with_item_ > 1 Then
               Exit Function
             End If
             If squre_distance_point_point(m_poi(item0(it%).data(0).poi(0)).data(0).data0.coordinate, _
                  m_poi(item0(it%).data(0).poi(1)).data(0).data0.coordinate) < _
                   squre_distance_point_point(m_poi(item0(it%).data(0).poi(2)).data(0).data0.coordinate, _
                       m_poi(item0(it%).data(0).poi(3)).data(0).data0.coordinate) Then
                item0(m(1)).data(0).big_or_smamll = True
             End If
      combine_line_value_with_item_ = set_general_string(m(0), m(1), 0, 0, _
        "-1", line_value(lv%).data(0).data0.value, "0", "0", item0(it%).data(0).value, 0, 0, 0, _
          temp_record, 0, 0)
            If combine_line_value_with_item_ > 1 Then
              Exit Function
            End If
     ElseIf item0(it%).data(0).diff_type = 4 Or item0(it%).data(0).diff_type = 8 Then
      combine_line_value_with_item_ = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
             item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), "*", _
              item0(it%).data(0).n(0), item0(it%).data(0).n(1), item0(it%).data(0).n(0), _
               item0(it%).data(0).n(1), item0(it%).data(0).line_no(0), item0(it%).data(0).line_no(0), _
                "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, m(0), _
                  no_reduce, 0, condition_data0, False)
             If combine_line_value_with_item_ > 1 Then
               Exit Function
             End If
      combine_line_value_with_item_ = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
             0, 0, "~", item0(it%).data(0).n(0), item0(it%).data(0).n(1), _
              0, 0, item0(it%).data(0).line_no(0), 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, _
               0, m(1), no_reduce, 0, condition_data0, False)
             If combine_line_value_with_item_ > 1 Then
               Exit Function
             End If
       combine_line_value_with_item_ = set_general_string(m(0), m(1), 0, 0, _
        "1", time_string("-1", line_value(lv%).data(0).data0.value, True, False), "0", "0", _
          item0(it%).data(0).value, 0, 0, 0, temp_record, 0, 0)
           If combine_line_value_with_item_ > 1 Then
            Exit Function
           End If
      ElseIf item0(it%).data(0).diff_type = 7 Or item0(it%).data(0).diff_type = 6 Then
       combine_line_value_with_item_ = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
             item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), "*", _
              item0(it%).data(0).n(0), item0(it%).data(0).n(1), item0(it%).data(0).n(0), _
               item0(it%).data(0).n(1), item0(it%).data(0).line_no(0), item0(it%).data(0).line_no(0), _
                "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, m(0), _
                  no_reduce, 0, condition_data0, False)
             If combine_line_value_with_item_ > 1 Then
               Exit Function
             End If
       combine_line_value_with_item_ = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
             0, 0, "~", item0(it%).data(0).n(0), item0(it%).data(0).n(1), _
              0, 0, item0(it%).data(0).line_no(0), 0, "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, _
                0, m(1), no_reduce, 0, condition_data0, False)
             If combine_line_value_with_item_ > 1 Then
               Exit Function
             End If
        combine_line_value_with_item_ = set_general_string(m(0), m(1), 0, 0, _
         "1", line_value(lv%).data(0).data0.value, "0", "0", item0(it%).data(0).value, 0, 0, 0, _
           temp_record, 0, 0)
            If combine_line_value_with_item_ > 1 Then
             Exit Function
            End If
      End If
     End If
   Else 'if value=""
     If j% = 0 Then
      temp_record1 = temp_record_data.data0.condition_data
       combine_line_value_with_item_ = set_item0(item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), 0, 0, _
          "~", item0(it%).data(0).n(2), item0(it%).data(0).n(3), 0, 0, item0(it%).data(0).line_no(1), _
           0, "1", "1", line_value(lv%).data(0).data0.value, "", "1", 0, record_data0.data0.condition_data, _
             it%, 0, no_reduce, 0, condition_data0, False)
             If combine_line_value_with_item_ > 1 Then
               Exit Function
             End If
     ElseIf j% = 1 Then
      combine_line_value_with_item_ = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), 0, 0, _
          "~", item0(it%).data(0).n(0), item0(it%).data(0).n(1), 0, 0, item0(it%).data(0).line_no(0), _
           0, "1", "1", line_value(lv%).data(0).data0.value, "", "1", 0, record_data0.data0.condition_data, _
             it%, 0, no_reduce, 0, condition_data0, False)
             If combine_line_value_with_item_ > 1 Then
               Exit Function
             End If
     ElseIf j% = 2 Then
     temp_record1.condition_no = 1
      temp_record1.condition(1).no = lv%
       temp_record1.condition(1).ty = line_value_
        'temp_record1.para
      If item0(it%).data(0).n(1) = item0(it%).data(0).n(2) And _
          item0(it%).data(0).n(0) = item0(it%).data(0).n(4) And _
           item0(it%).data(0).n(3) = item0(it%).data(0).n(5) Then
            Call set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
               item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), "*", item0(it%).data(0).n(0), item0(it%).data(0).n(1), _
                item0(it%).data(0).n(0), item0(it%).data(0).n(1), item0(it%).data(0).line_no(0), item0(it%).data(0).line_no(0), _
                 "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, it_(0), 0, 0, condition_data0, False)
            Call set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
               0, 0, "~", item0(it%).data(0).n(0), item0(it%).data(0).n(1), _
                0, 0, item0(it%).data(0).line_no(0), 0, _
                 "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, it_(1), 0, 0, condition_data0, False)
           combine_line_value_with_item_ = add_new_item_to_item(it_(0), it_(1), "-1", _
             line_value(lv%).data(0).data0.value_, it%, temp_record1)
           If combine_line_value_with_item_ > 1 Then
            Exit Function
           End If
            Call set_item0(item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), _
               item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), "*", item0(it%).data(0).n(2), item0(it%).data(0).n(3), _
                item0(it%).data(0).n(2), item0(it%).data(0).n(3), item0(it%).data(0).line_no(0), item0(it%).data(0).line_no(0), _
                 "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, it_(0), 0, 0, condition_data0, False)
            Call set_item0(item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), _
               0, 0, "~", item0(it%).data(0).n(2), item0(it%).data(0).n(3), _
                0, 0, item0(it%).data(0).line_no(0), 0, _
                 "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, it_(1), 0, 0, condition_data0, False)
           combine_line_value_with_item_ = add_new_item_to_item(it_(0), it_(1), "-1", _
             line_value(lv%).data(0).data0.value_, it%, temp_record1)
           If combine_line_value_with_item_ > 1 Then
            Exit Function
           End If
      ElseIf item0(it%).data(0).n(0) = item0(it%).data(0).n(2) And _
         item0(it%).data(0).n(1) = item0(it%).data(0).n(4) And _
              item0(it%).data(0).n(3) = item0(it%).data(0).n(5) Then
             Call set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
               item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), "*", item0(it%).data(0).n(0), item0(it%).data(0).n(1), _
                item0(it%).data(0).n(0), item0(it%).data(0).n(1), item0(it%).data(0).line_no(0), item0(it%).data(0).line_no(0), _
                 "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, it_(0), 0, 0, condition_data0, False)
            Call set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
               0, 0, "~", item0(it%).data(0).n(0), item0(it%).data(0).n(1), _
                0, 0, item0(it%).data(0).line_no(0), 0, _
                 "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, it_(1), 0, 0, condition_data0, False)
           combine_line_value_with_item_ = add_new_item_to_item(it_(0), it_(1), "1", _
             line_value(lv%).data(0).data0.value_, it%, temp_record1)
           If combine_line_value_with_item_ > 1 Then
            Exit Function
           End If
     ElseIf item0(it%).data(0).n(1) = item0(it%).data(0).n(3) And _
          item0(it%).data(0).n(0) = item0(it%).data(0).n(4) And _
              item0(it%).data(0).n(2) = item0(it%).data(0).n(5) Then
            Call set_item0(item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), _
               item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), "*", item0(it%).data(0).n(2), item0(it%).data(0).n(3), _
                item0(it%).data(0).n(2), item0(it%).data(0).n(3), item0(it%).data(0).line_no(0), item0(it%).data(0).line_no(0), _
                 "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, it_(0), 0, 0, condition_data0, False)
            Call set_item0(item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), _
               0, 0, "~", item0(it%).data(0).n(2), item0(it%).data(0).n(3), _
                0, 0, item0(it%).data(0).line_no(0), 0, _
                 "1", "1", "1", "", "1", 0, record_data0.data0.condition_data, 0, it_(1), 0, 0, condition_data0, False)
           combine_line_value_with_item_ = add_new_item_to_item(it_(0), it_(1), "1", _
             line_value(lv%).data(0).data0.value_, it%, temp_record1)
           If combine_line_value_with_item_ > 1 Then
            Exit Function
           End If
         End If
     End If
    End If
    End If
   ElseIf item0(it%).data(0).sig = "/" Then
    If item0(it%).data(0).value <> "" Then
           temp_record.record_data = temp_record_data
     Call add_record_to_record(item0(it%).data(0).record_for_value.data0.condition_data, _
                                               temp_record.record_data.data0.condition_data)
       If j% = 0 Then
         combine_line_value_with_item_ = set_element_value(it%, 1, _
           divide_string(line_value(lv%).data(0).data0.value, item0(it%).data(0).value, True, False), _
            temp_record, no_reduce)
             If combine_line_value_with_item_ > 1 Then
               Exit Function
             End If
       ElseIf j% = 1 Then
         combine_line_value_with_item_ = set_element_value(it%, 0, _
           time_string(line_value(lv%).data(0).data0.value, item0(it%).data(0).value, _
            True, False), temp_record, no_reduce)
             If combine_line_value_with_item_ > 1 Then
               Exit Function
             End If
       Else
         temp_record.record_data = temp_record_data
           Call add_record_to_record(item0(it%).data(0).record_for_value.data0.condition_data, _
                     temp_record.record_data.data0.condition_data)
        If item0(it%).data(0).diff_type = 3 Or item0(it%).data(0).diff_type = 5 Then
             combine_line_value_with_item_ = set_line_value(item0(it%).data(0).poi(2), _
               item0(it%).data(0).poi(3), divide_string(line_value(lv%).data(0).data0.value, _
                 add_string("1", item0(it%).data(0).value, False, False), True, False), _
                   item0(it%).data(0).n(2), item0(it%).data(0).n(3), item0(it%).data(0).line_no(1), _
                     temp_record, 0, 0, False)
            If combine_line_value_with_item_ > 1 Then
             Exit Function
            End If
         ElseIf item0(it%).data(0).diff_type = 4 Or item0(it%).data(0).diff_type = 8 Then
             combine_line_value_with_item_ = set_line_value(item0(it%).data(0).poi(2), _
                 item0(it%).data(0).poi(3), divide_string(line_value(lv%).data(0).data0.value, _
                    minus_string(item0(it%).data(0).value, "1", False, False), True, False), _
                       item0(it%).data(0).n(2), item0(it%).data(0).n(3), item0(it%).data(0).line_no(1), _
                          temp_record, 0, 0, False)
            If combine_line_value_with_item_ > 1 Then
             Exit Function
            End If
          ElseIf item0(it%).data(0).diff_type = 7 Or item0(it%).data(0).diff_type = 6 Then
              combine_line_value_with_item_ = set_line_value(item0(it%).data(0).poi(2), _
                item0(it%).data(0).poi(3), divide_string(line_value(lv%).data(0).data0.value, _
                   minus_string("1", item0(it%).data(0).value, False, False), True, False), _
                     item0(it%).data(0).n(2), item0(it%).data(0).n(3), item0(it%).data(0).line_no(1), _
                       temp_record, 0, 0, False)
             If combine_line_value_with_item_ > 1 Then
               Exit Function
             End If
           End If
         End If
    Else
          temp_record1 = temp_record_data.data0.condition_data
      If j% = 0 Then
        combine_line_value_with_item_ = set_item0(0, 0, item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), _
         "/", 0, 0, item0(it%).data(0).n(2), item0(it%).data(0).n(3), 0, item0(it%).data(0).line_no(1), _
           "1", "1", line_value(lv%).data(0).data0.value, "", "1", 0, record_data0.data0.condition_data, _
              it%, 0, no_reduce, 0, condition_data0, False)
              If combine_line_value_with_item_ > 1 Then
               Exit Function
             End If
     ElseIf j% = 1 Then
        combine_line_value_with_item_ = set_item0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
          0, 0, "~", item0(it%).data(0).n(0), item0(it%).data(0).n(1), 0, 0, item0(it%).data(0).line_no(1), 0, _
            "1", "1", divide_string("1", line_value(lv%).data(0).data0.value, True, False), "", "1", 0, _
               record_data0.data0.condition_data, it%, 0, no_reduce, 0, condition_data0, False)
             If combine_line_value_with_item_ > 1 Then
               Exit Function
             End If
      End If
    End If
   ElseIf item0(it%).data(0).sig = "~" Then
    If item0(it%).data(0).value = "" Then
     item0(it%).data(0).value = line_value(lv%).data(0).data0.value
      item0(it%).data(0).record_for_value.data0.condition_data = temp_record_data.data0.condition_data
           combine_line_value_with_item_ = combine_item_with_general_string_(it%, -1)
    End If
   End If
  End If
 End If
End Function

Public Function combine_two_total_angle(ByVal t_A1%, ByVal t_A2%, out_t_A%, ty As Byte, last As Byte) As Boolean
Dim t_A(2) As total_angle_data_type
Dim n%, ty1%
If t_A1% = 0 Or t_A2% = 0 Then
 ty = 0
ElseIf t_A1% = 0 Or t_A2% = 0 Then
 ty = 1
  combine_two_total_angle = True
Else
t_A(0) = T_angle(t_A1%).data(0)
t_A(1) = T_angle(t_A2%).data(0)
If t_A(0).line_no(0) = t_A(1).line_no(0) Then
 If t_A(0).line_no(1) < t_A(1).line_no(1) Then
   t_A(2).line_no(2) = t_A(0).line_no(1)
    t_A(2).line_no(3) = t_A(1).line_no(1)
 Else
   t_A(2).line_no(2) = t_A(1).line_no(1)
    t_A(2).line_no(3) = t_A(0).line_no(1)
 End If
 ty1% = 0
ElseIf t_A(0).line_no(1) = t_A(1).line_no(1) Then
 If t_A(0).line_no(0) < t_A(0).line_no(0) Then
  t_A(2).line_no(2) = t_A(0).line_no(0)
   t_A(2).line_no(3) = t_A(1).line_no(0)
 Else
  t_A(2).line_no(2) = t_A(1).line_no(0)
   t_A(2).line_no(3) = t_A(0).line_no(0)
 End If
 ty1% = 1
ElseIf t_A(0).line_no(0) = t_A(1).line_no(1) Then
 If t_A(0).line_no(1) < t_A(0).line_no(0) Then
  t_A(2).line_no(2) = t_A(0).line_no(1)
   t_A(2).line_no(3) = t_A(1).line_no(0)
 Else
  t_A(2).line_no(2) = t_A(1).line_no(0)
   t_A(2).line_no(3) = t_A(0).line_no(1)
 End If
 ty1% = 2
ElseIf t_A(0).line_no(1) = t_A(1).line_no(0) Then
 If t_A(0).line_no(0) < t_A(0).line_no(1) Then
  t_A(2).line_no(2) = t_A(0).line_no(0)
   t_A(2).line_no(3) = t_A(1).line_no(1)
 Else
  t_A(2).line_no(2) = t_A(1).line_no(1)
   t_A(2).line_no(3) = t_A(0).line_no(0)
 End If
 ty1% = 3
Else
 Exit Function
End If
combine_two_total_angle = True
If search_for_total_angle(t_A(2), n%, 0, 0) Then
 out_t_A% = n%
Else
 Call set_total_angle(t_A(2), n%, n%, False)
End If
out_t_A% = n%
If t_A1% < t_A2% Then
 If n% < t_A2% Then
  last = 1
 Else
  last = 2
 End If
Else
 If n% < t_A1% Then
  last = 0
 Else
  last = 2
 End If
End If
If ty1% = 0 Then
 If t_A(2).line_no(0) = t_A(0).line_no(1) Then
 ty = 7
 Else
 ty = 8
 End If
ElseIf ty1% = 1 Then
 If t_A(2).line_no(0) = t_A(0).line_no(0) Then
  ty = 4
 Else
  ty = 6
 End If
ElseIf ty1% = 2 Then
 If t_A(2).line_no(0) = t_A(0).line_no(1) Then
  ty = 5
 Else
  ty = 10
 End If
ElseIf ty1% = 3 Then
 If t_A(2).line_no(0) = t_A(0).line_no(0) Then
  ty = 3
 Else
  ty = 9
 End If
End If
End If
End Function
Public Function combine_two_total_angle_with_para(tA1%, tA2%, para1$, para2$, v$) As Byte
Dim tA3% ' 按系数合并角, 不能完全合并,排除最后的角,
Dim ty As Byte
Dim last As Byte
v$ = "0"
If tA1% = tA2% Then
 para1$ = add_string(para1$, para2$, True, False)
  If para1$ = "0" Then
   tA1% = 0
  End If
  para2$ = "0"
   tA2% = 0
          combine_two_total_angle_with_para = True
Else
If combine_two_total_angle(tA1%, tA2%, tA3%, ty, last) Then '无系数合并
 If ty = 3 Or ty = 5 Then
  If para1$ = para2$ Then
   para2$ = "0"
    tA2% = 0
     tA1% = tA3%
          combine_two_total_angle_with_para = True
  Else
   If last = 0 Then
    tA1% = tA3%
     para2$ = minus_string(para2$, para1$, True, False)
   ElseIf last = 1 Then
    tA2% = tA3%
     para1$ = minus_string(para1$, para2$, True, False)
   End If
  End If
 ElseIf ty = 4 Or ty = 8 Then
  If para1$ = time_string(para2$, "-1", True, False) Then
   para2$ = "0"
    tA2% = 0
     tA1% = tA3%
          combine_two_total_angle_with_para = True
 Else
   If last = 0 Then
    tA1% = tA3%
     para2$ = add_string(para2$, para1$, True, False)
   ElseIf last = 1 Then
    tA2% = tA3%
     para1$ = add_string(para2$, para1$, False, False)
      para2 = time_string("-1", para1, True, False)
   End If
  End If
 ElseIf ty = 6 Or ty = 7 Then
  If para1$ = time_string(para2$, "-1", True, False) Then
   para1$ = "0"
    tA1% = 0
     tA2% = tA3%
          combine_two_total_angle_with_para = True
  Else
   If last = 0 Then
    tA1% = tA3%
     para2$ = add_string(para2$, para1$, True, False)
      para1$ = time_string("-1", para1$, True, False)
   ElseIf last = 1 Then
    tA2% = tA3%
     para1$ = add_string(para1$, para2$, True, False)
   End If
  End If
  ElseIf ty = 9 Or ty = 10 Then
    If para1$ = para2$ Then
       v$ = time_string(para1$, "360", True, False)
        para1$ = time_string("-1", para1$, True, False)
         para2$ = time_string("-1", para2$, True, False)
          combine_two_total_angle_with_para = True
    Else
     If last = 0 Then
      tA1% = tA3%
       para2$ = minus_string(para2$, para1$, True, False)
        v$ = time_string(para1$, "360", True, False)
         para1$ = time_string("-1", para1$, True, False)
     ElseIf last = 1 Then
      tA2% = tA3%
       para1$ = minus_string(para1$, para2$, True, False)
        v$ = time_string(para2$, "360", True, False)
         para2$ = time_string("-1", para2$, True, False)
     End If
    End If
  End If
End If
End If
End Function

Public Function combine_two_t_angle3_value0(T_A3_v1 As angle3_value_data0_type, _
               T_A3_v2 As angle3_value_data0_type, t_angle3_v As angle3_value_data0_type) As Boolean
Dim i%, j%
For i% = 0 To 2
 For j% = 0 To 2
  If combine_two_t_angle3_value_(T_A3_v1, _
         T_A3_v1, i%, j%, t_angle3_v) Then
   combine_two_t_angle3_value0 = True
    Exit Function
  End If
 Next j%
Next i%
End Function

Public Function combine_two_t_angle3_value_(tA3_v1 As angle3_value_data0_type, tA3_v2 As angle3_value_data0_type, _
                k1%, k2%, t_A3_v3 As angle3_value_data0_type) As Boolean
Dim tA%, i%, j%
Dim n1(2) As Integer
Dim n2(2) As Integer
Dim ty As Byte
Dim A(1, 3) As Integer
Dim p(1, 3) As String
Dim ts$
Dim last As Byte
Dim angle_no As Integer
angle_no = 6
For j% = 0 To 2
A(0, j%) = tA3_v1.angle((k1% + j%) Mod 3)
If A(0, j%) = 0 Then
 angle_no = angle_no - 1
End If
p(0, j%) = tA3_v1.para((k1% + j%) Mod 3)
A(1, j%) = tA3_v2.angle((k2% + j%) Mod 3)
If A(1, j%) = 0 Then
 angle_no = angle_no - 1
End If
p(1, j%) = tA3_v2.para((k2% + j%) Mod 3)
Next j%
p(0, 3) = tA3_v1.value
p(1, 3) = tA3_v2.value
 If combine_two_total_angle(A(0, 0), A(1, 0), tA%, ty, last) Then
  If ty = 1 Then
     Call cal_eight_para(p(), 1)
     p(0, 0) = "0"
      A(0, 0) = 0
       A(1, 0) = 0
        angle_no = angle_no - 2
   ElseIf ty = 3 Or ty = 5 Then
     Call cal_eight_para(p(), 0)
     A(0, 0) = tA%
      A(1, 0) = "0"
        angle_no = angle_no - 1
  ElseIf ty = 4 Or ty = 8 Then
     Call cal_eight_para(p(), 1)
      A(0, 0) = tA%
       A(1, 0) = "0"
         angle_no = angle_no - 1
 ElseIf ty = 6 Or ty = 7 Then
     Call cal_eight_para(p(), 2)
      A(0, 0) = tA%
       A(1, 0) = "0"
        angle_no = angle_no - 1
 ElseIf ty = 9 Or ty = 10 Then
     Call cal_eight_para(p(), 0)
     A(0, 0) = tA%
      A(1, 0) = "0"
         angle_no = angle_no - 1
      p(0, 0) = time_string(p(0, 0), "-1", True, True)
        p(0, 3) = add_string(p(0, 3), time_string(p(0, 0), "360", False, True), True, True)
  Else
   Exit Function
  End If
Else
   Exit Function
End If
 For i% = 0 To 1
  For j% = 0 To 2
   If p(i%, j%) = "0" And A(i%, j%) > 0 Then
    A(i%, j%) = 0
        angle_no = angle_no - 1
   End If
  Next j%
 Next i%
If angle_no > 3 Then
For i% = 1 To 2
 For j% = 1 To 2
  If combine_two_total_angle_with_para(A(0, i%), A(1, j%), p(0, i%), p(1, j%), ts$) Then
   p(0, 3) = minus_string(p(0, 3), ts$, True, False)
   If i% = 2 Then
    Call exchange_two_integer(A(0, 1), A(0, 2))
    Call exchange_string(p(0, 1), p(0, 2))
   End If
   If j% = 2 Then
    Call exchange_two_integer(A(1, 1), A(1, 2))
    Call exchange_string(p(1, 1), p(1, 2))
   End If
   GoTo combine_two_t_angle3_value_mark0
   End If
 Next j%
Next i%
End If
combine_two_t_angle3_value_mark0:
If angle_no > 3 Then
 If combine_two_total_angle_with_para(A(0, 2), A(1, 2), p(0, 2), p(1, 2), ts$) Then
   p(0, 3) = minus_string(p(0, 3), ts$, True, False)
 End If
End If
If angle_no < 4 Then
  angle_no = 0
   t_A3_v3.value = add_string(p(0, 3), p(1, 3), True, False)
    For i% = 0 To 2
     If p(0, i%) <> "0" Then
      t_A3_v3.para(angle_no) = p(0, i%)
      t_A3_v3.angle(angle_no) = A(0, i%)
     End If
    Next i%
    For i% = 0 To 2
     If p(1, i%) <> "0" Then
      t_A3_v3.para(angle_no) = p(1, i%)
      t_A3_v3.angle(angle_no) = A(1, i%)
     End If
    Next i%
    For i% = angle_no To 2
      t_A3_v3.para(angle_no) = "0"
      t_A3_v3.angle(angle_no) = 0
    Next i%
    combine_two_t_angle3_value_ = True
End If
End Function

Public Sub cal_eight_para(p() As String, minus_or_add As Byte)
Dim tp1$
Dim tp2$
tp1$ = p(0, 0)
tp2$ = p(1, 0)
 p(0, 0) = time_string(p(0, 0), tp2$, True, False)
 p(0, 1) = time_string(p(0, 1), tp2$, True, False)
 p(0, 2) = time_string(p(0, 2), tp2$, True, False)
 p(0, 2) = time_string(p(0, 3), tp2$, True, False)
 p(1, 0) = time_string(p(1, 0), tp1$, True, False)
 p(1, 1) = time_string(p(1, 1), tp1$, True, False)
 p(1, 2) = time_string(p(1, 2), tp1$, True, False)
 p(1, 2) = time_string(p(1, 3), tp1$, True, False)
 If minus_or_add = 0 Then
  p(1, 0) = "0"
 ElseIf minus_or_add = 1 Then ' 0-1
  p(1, 0) = "0"
  p(1, 1) = time_string(p(1, 1), "-1", True, False)
  p(1, 2) = time_string(p(1, 2), "-1", True, False)
  p(1, 2) = time_string(p(1, 3), "-1", True, False)
 ElseIf minus_or_add = 2 Then
  p(0, 0) = "0"
  p(0, 1) = time_string(p(0, 1), "-1", True, False)
  p(0, 2) = time_string(p(0, 2), "-1", True, False)
  p(0, 2) = time_string(p(0, 3), "-1", True, False)
 End If
End Sub

Public Sub combine_two_t_angle3_value(t_A3_v%)
Dim t_angle3_v As angle3_value_data0_type
Dim i%, j%
Dim n(2), m(2) As Integer
For i% = 0 To 2
  n(0) = i%
  n(1) = (i% + 1) Mod 3
  n(2) = (i% + 2) Mod 3
 For j% = 0 To 2
  m(0) = j%
  m(1) = (j% + 1) Mod 3
  m(2) = (j% + 2) Mod 3
  
 Next j%
Next i%
End Sub


Public Function read_angle3_value(a3_v1 As angle3_value_data0_type, k%, _
      A3_v2 As angle3_value_data0_type, A3_v3 As angle3_value_data0_type) As Boolean
        'ty新类型
Dim n(2) As Integer
Dim k_%
k_% = k% - 3
n(0) = k_%
n(1) = (k_% + 1) Mod 3
n(2) = (k_% + 2) Mod 3
A3_v2.no_zero_angle = a3_v1.no_zero_angle
A3_v3.no_zero_angle = a3_v1.no_zero_angle
'记录第三角,和值
A3_v2.angle(2) = a3_v1.angle(n(2))
A3_v2.para(2) = a3_v1.para(n(2))
A3_v2.value = a3_v1.value
A3_v3.angle(2) = a3_v1.angle(n(2))
A3_v3.para(2) = a3_v1.para(n(2))
A3_v3.value = a3_v1.value
'******************
'设置和角
A3_v2.angle(0) = a3_v1.angle(k%)
A3_v2.para(0) = a3_v1.para(n(0))
A3_v3.angle(0) = a3_v1.angle(k%)
A3_v3.para(0) = a3_v1.para(n(1))
If a3_v1.ty(k_%) = 3 Or a3_v1.ty(k_%) = 5 Then
A3_v2.para(1) = minus_string(a3_v1.para(n(1)), a3_v1.para(n(0)), True, False)
If A3_v2.para(1) <> "0" Then
 A3_v2.angle(1) = a3_v1.angle(n(1))
End If
A3_v3.para(1) = minus_string(a3_v1.para(n(0)), a3_v1.para(n(1)), True, False)
If A3_v3.para(1) <> "0" Then
 A3_v3.angle(1) = a3_v1.angle(n(0))
End If
read_angle3_value = True
'ElseIf A3_v1.ty(k%) = 4 Then

'ElseIf A3_v1.ty(k%) = 6 Then
'ElseIf A3_v1.ty(k%) = 7 Then
'ElseIf A3_v1.ty(k%) = 8 Then
End If
End Function
Public Function combine_two_relation_on_line(ByVal re_on_l%) As Byte
Dim tn%, j%, k%
Dim l(3) As Integer
Dim p(2) As Integer
Dim q(2) As Integer
Dim tp1(1) As Integer
Dim tp2(1) As Integer
Dim re_v1(1) As String
Dim re_v2(1) As String
Dim temp_record As total_record_type
If th_chose(-4).chose = 0 Or th_chose(-6).chose = 0 Then
'未选中美耐劳斯定理
 Exit Function
End If
 For tn% = 1 To re_on_l% - 1
  If relation_on_line(re_on_l%).data(0).data0.line_no <> relation_on_line(tn%).data(0).data0.line_no Then
   For j% = 0 To 2
    For k% = 0 To 2
     If relation_on_line(re_on_l%).data(0).data0.poi(j%) = relation_on_line(tn%).data(0).data0.poi(k%) Then
      '两条直线上的比,有交点
      GoTo combine_two_relation_on_line_mark0
     End If
    Next k%
   Next j%
   '没有交点
      GoTo combine_two_relation_on_line_mark1
'********************************************
combine_two_relation_on_line_mark0:
temp_record.record_data.data0.theorem_no = -4 '美耐劳斯定理
'有交点,交点是relation_on_line(re_on_l%).data(0).data0.poi(j%)和
'relation_on_line(tn%).data(0).data0.poi(k%)引用美耐劳斯定理
temp_record.record_data.data0.condition_data.condition_no = 0
Call add_record_to_record(relation_on_line(re_on_l%).data(0).record, temp_record.record_data.data0.condition_data)
Call add_record_to_record(relation_on_line(tn%).data(0).record, temp_record.record_data.data0.condition_data)
'设置两线和五个点
l(0) = relation_on_line(re_on_l%).data(0).data0.line_no
l(1) = relation_on_line(tn%).data(0).data0.line_no
q(2) = relation_on_line(re_on_l%).data(0).data0.poi(j%) '交点
If j% = 0 Then
 tp1(0) = relation_on_line(re_on_l%).data(0).data0.poi(1) '另两点
 tp1(1) = relation_on_line(re_on_l%).data(0).data0.poi(2)
 'q(2)tp1(0)/tp1(0)tp1(1)
 'q(2)tp1(1)/tp1(0)tp1(1)
 re_v1(0) = relation_on_line(re_on_l%).data(0).data0.value
 re_v1(1) = add_string("1", relation_on_line(re_on_l%).data(0).data0.value, True, False)
ElseIf j% = 1 Then
 tp1(0) = relation_on_line(re_on_l%).data(0).data0.poi(2)
 tp1(1) = relation_on_line(re_on_l%).data(0).data0.poi(0)
 re_v1(0) = divide_string("1", add_string("1", _
      relation_on_line(re_on_l%).data(0).data0.value, False, False), True, False)
 re_v1(1) = divide_string(relation_on_line(re_on_l%).data(0).data0.value, add_string("1", _
      relation_on_line(re_on_l%).data(0).data0.value, False, False), True, False)
Else
 tp1(0) = relation_on_line(re_on_l%).data(0).data0.poi(0)
 tp1(1) = relation_on_line(re_on_l%).data(0).data0.poi(1)
  re_v1(0) = divide_string(add_string("1", relation_on_line(re_on_l%).data(0).data0.value, False, False), _
       relation_on_line(re_on_l%).data(0).data0.value, True, False)
  re_v1(1) = divide_string("1", relation_on_line(re_on_l%).data(0).data0.value, True, False)
End If
If k% = 0 Then
 tp2(0) = relation_on_line(tn%).data(0).data0.poi(1) '另两点
 tp2(1) = relation_on_line(tn%).data(0).data0.poi(2)
  re_v2(0) = divide_string("1", relation_on_line(tn%).data(0).data0.value, True, False)
  re_v2(1) = divide_string("1", add_string("1", _
      relation_on_line(tn%).data(0).data0.value, False, False), True, False)
ElseIf k% = 1 Then
 tp2(0) = relation_on_line(tn%).data(0).data0.poi(2)
 tp2(1) = relation_on_line(tn%).data(0).data0.poi(0)
 re_v2(0) = add_string("1", relation_on_line(tn%).data(0).data0.value, True, False)
 re_v2(1) = divide_string(add_string("1", relation_on_line(tn%).data(0).data0.value, False, False), _
       relation_on_line(tn%).data(0).data0.value, True, False)
Else
 tp2(0) = relation_on_line(tn%).data(0).data0.poi(0)
 tp2(1) = relation_on_line(tn%).data(0).data0.poi(1)
 re_v2(0) = divide_string(relation_on_line(tn%).data(0).data0.value, add_string("1", _
      relation_on_line(tn%).data(0).data0.value, False, False), True, False)
 re_v2(1) = relation_on_line(tn%).data(0).data0.value
End If
If th_chose(-4).chose = 1 Then
temp_record.record_data.data0.theorem_no = -4 '美耐劳斯定理
 If j% = 0 Or k% = 2 Or j% = 2 Or k% = 0 Then
  l(2) = line_number0(tp1(0), tp2(1), 0, 0)  '
  l(3) = line_number0(tp1(1), tp2(0), 0, 0)  '
  p(0) = is_line_line_intersect(l(2), l(3), 0, 0, False)
   If p(0) > 0 Then
    l(2) = line_number0(tp1(1), tp2(1), 0, 0)  '
    l(3) = line_number0(q(2), p(0), 0, 0)  '
    p(1) = is_line_line_intersect(l(2), l(3), 0, 0, False)
     If p(1) > 0 Then
      combine_two_relation_on_line = set_Drelation(tp2(1), p(1), p(1), tp1(1), _
          0, 0, 0, 0, 0, 0, time_string(re_v1(0), re_v2(0), True, False), temp_record, _
           0, 0, 0, 0, 0, False)
       If combine_two_relation_on_line > 1 Then
          Exit Function
       End If
     End If
   End If
 End If
End If
If th_chose(-6).chose = 1 Then
temp_record.record_data.data0.theorem_no = -6 '美耐劳斯定理
l(2) = line_number0(tp1(0), tp2(0), 0, 0)  '
l(3) = line_number0(tp1(1), tp2(1), 0, 0)  '
'
If l(2) = l(3) Then
 GoTo combine_two_relation_on_line_mark2
End If
p(2) = is_line_line_intersect(l(2), l(3), 0, 0, False)
If p(2) > 0 Then
combine_two_relation_on_line = set_Drelation(p(2), tp2(1), p(2), tp1(1), 0, 0, 0, 0, _
     0, 0, time_string(re_v1(0), re_v2(0), True, False), temp_record, 0, 0, 0, 0, 0, False)
If combine_two_relation_on_line > 1 Then
   Exit Function
End If
combine_two_relation_on_line = set_Drelation(p(2), tp2(0), p(2), tp1(0), 0, 0, 0, 0, _
     0, 0, time_string(re_v1(1), re_v2(1), True, False), temp_record, 0, 0, 0, 0, 0, False)
If combine_two_relation_on_line > 1 Then
   Exit Function
End If
End If
combine_two_relation_on_line_mark2:
l(2) = line_number0(tp1(0), tp2(1), 0, 0)
l(3) = line_number0(tp1(1), tp2(0), 0, 0)
If l(2) = l(3) Then
 GoTo combine_two_relation_on_line_mark1
End If
p(2) = is_line_line_intersect(l(2), l(3), 0, 0, False)
If p(2) > 0 Then
combine_two_relation_on_line = set_Drelation(p(2), tp2(0), p(2), tp1(1), 0, 0, 0, 0, _
     0, 0, time_string(re_v1(0), re_v2(1), True, False), temp_record, 0, 0, 0, 0, 0, False)
If combine_two_relation_on_line > 1 Then
   Exit Function
End If
combine_two_relation_on_line = set_Drelation(p(2), tp2(1), p(2), tp1(0), 0, 0, 0, 0, _
     0, 0, time_string(re_v1(1), re_v2(0), True, False), temp_record, 0, 0, 0, 0, 0, False)
If combine_two_relation_on_line > 1 Then
   Exit Function
End If
End If
End If
combine_two_relation_on_line_mark1:
End If
 Next tn%
End Function

Public Function combine_two_item_for_p(ByVal i1%, ByVal i2%, ByVal p1$, ByVal p2$, _
                   outI1%, outI2%, outp1$, outp2$, c_data As condition_data_type) As Byte
'两平方项合并,p勾股定理
Dim n%, n1%, it0%
Dim l(2) As Integer
Dim ty As Byte
Dim dir As String
Dim tn(3) As Integer
Dim tp(3) As Integer
Dim re As condition_data_type
If th_chose(51).chose = 0 Then
 Exit Function
End If
dir = "1"
If item0(i1%).data(0).sig = "*" And item0(i2%).data(0).sig = "*" Then
 If item0(i1%).data(0).poi(1) > 0 And item0(i1%).data(0).poi(0) = item0(i1%).data(0).poi(2) And _
      item0(i1%).data(0).poi(1) = item0(i1%).data(0).poi(3) Then '平方
   If item0(i2%).data(0).poi(1) > 0 And item0(i2%).data(0).poi(0) = item0(i2%).data(0).poi(2) And _
      item0(i2%).data(0).poi(1) = item0(i2%).data(0).poi(3) Then '平方
       If p1$ = p2$ Then '相等系数
        ty = 0 '两直角边
       ElseIf p1$ = time_string(p2$, "-1", True, False) Then
        ty = 1 '相反系数
       Else
        Exit Function
       End If
     If item0(i1%).data(0).line_no(0) = item0(i2%).data(0).line_no(0) And ty = 1 Then '平方差
        If item0(i1%).data(0).poi(0) = item0(i2%).data(0).poi(0) Then
         'Call set_item0(item0(i1%).data(0).poi(0), _
                item0(i1%).data(0).poi(1), item0(i2%).data(0).poi(0), _
                 item0(i2%).data(0).poi(1), "+", item0(i1%).data(0).n(0), _
                  item0(i1%).data(0).n(1), item0(i2%).data(0).n(0), _
                   item0(i2%).data(0).n(1), item0(i1%).data(0).line_no(0), _
                    item0(i1%).data(0).line_no(0), "1", "1", "1", "", "1", 0, _
                     re, 0, it0%, 0)
         
        If item0(i1%).data(0).n(1) > item0(i2%).data(0).n(1) Then '差
          tp(1) = item0(i1%).data(0).poi(1)
          tp(0) = item0(i2%).data(0).poi(1)
          Call set_item0(tp(0), tp(1), tp(0), tp(1), "*", 0, _
                  0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, _
                     re, 0, outI1%, 0, 0, condition_data0, False)
          Call set_item0(item0(i2%).data(0).poi(0), _
                item0(i2%).data(0).poi(1), tp(0), tp(1), "*", item0(i2%).data(0).n(0), _
                  item0(i2%).data(0).n(1), 0, 0, _
                    item0(i2%).data(0).line_no(0), _
                     0, "1", "1", "1", "", "1", 0, _
                      re, 0, outI2%, 0, 0, condition_data0, False)
            outp1$ = p1$
             outp2$ = time_string("2", p1$, True, False)

         Else
          tp(0) = item0(i1%).data(0).poi(1)
          tp(1) = item0(i2%).data(0).poi(1)
          Call set_item0(tp(0), tp(1), tp(0), tp(1), "*", 0, _
                  0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, _
                     re, 0, outI1%, 0, 0, condition_data0, False)
          Call set_item0(item0(i1%).data(0).poi(0), _
                item0(i1%).data(0).poi(1), tp(0), tp(1), "*", item0(i1%).data(0).n(0), _
                  item0(i1%).data(0).n(1), 0, 0, _
                    item0(i2%).data(0).line_no(0), _
                     0, "1", "1", "1", "", "1", 0, _
                      re, 0, outI2%, 0, 0, condition_data0, False)
            outp1$ = p2$
             outp2$ = time_string("2", p2$, True, False)
          dir = "-1"
         End If
     ElseIf item0(i1%).data(0).poi(0) = item0(i2%).data(0).poi(1) Then
         'Call set_item0(item0(i1%).data(0).poi(0), _
                item0(i1%).data(0).poi(1), item0(i2%).data(0).poi(0), _
                 item0(i2%).data(0).poi(1), "-", item0(i1%).data(0).n(0), _
                  item0(i1%).data(0).n(1), item0(i2%).data(0).n(0), _
                   item0(i2%).data(0).n(1), item0(i1%).data(0).line_no(0), _
                    item0(i1%).data(0).line_no(0), "1", "1", "1", "", "1", 0, _
                     re, 0, it0%, 0)
         tp(0) = item0(i2%).data(0).poi(0) '和
         tp(1) = item0(i1%).data(0).poi(1)
         Call set_item0(item0(i1%).data(0).poi(0), _
                item0(i1%).data(0).poi(1), tp(0), tp(1), "*", item0(i1%).data(0).n(0), _
                  item0(i1%).data(0).n(1), 0, 0, _
                    item0(i1%).data(0).line_no(0), _
                     0, "1", "1", "1", "", "1", 0, _
                     re, 0, outI1%, 0, 0, condition_data0, False)
         Call set_item0(item0(i2%).data(0).poi(0), _
                item0(i2%).data(0).poi(1), tp(0), tp(1), "*", item0(i2%).data(0).n(0), _
                  item0(i2%).data(0).n(1), 0, 0, _
                    item0(i2%).data(0).line_no(0), _
                     0, "1", "1", "1", "", "1", 0, _
                      re, 0, outI2%, 0, 0, condition_data0, False)
             If dir = "1" Then
               outp1$ = p1$
                outp2$ = p2$
              Else
               outp1$ = p2$
                outp2$ = p1$
              End If
     ElseIf item0(i1%).data(0).poi(1) = item0(i2%).data(0).poi(0) Then
         'Call set_item0(item0(i1%).data(0).poi(0), _
                item0(i1%).data(0).poi(1), item0(i2%).data(0).poi(0), _
                 item0(i2%).data(0).poi(1), "-", item0(i1%).data(0).n(0), _
                  item0(i1%).data(0).n(1), item0(i2%).data(0).n(0), _
                   item0(i2%).data(0).n(1), item0(i1%).data(0).line_no(0), _
                    item0(i1%).data(0).line_no(0), "1", "1", "1", "", "1", 0, _
                     re, 0, it0%, 0)
         tp(0) = item0(i1%).data(0).poi(0)
         tp(1) = item0(i2%).data(0).poi(1)
         Call set_item0(item0(i1%).data(0).poi(0), _
                item0(i1%).data(0).poi(1), tp(0), tp(1), "*", item0(i1%).data(0).n(0), _
                  item0(i1%).data(0).n(1), 0, 0, _
                    item0(i1%).data(0).line_no(0), _
                     0, "1", "1", "1", "", "1", 0, _
                     re, 0, outI1%, 0, 0, condition_data0, False)
         Call set_item0(item0(i2%).data(0).poi(0), _
                item0(i2%).data(0).poi(1), tp(0), tp(1), "*", item0(i2%).data(0).n(0), _
                  item0(i2%).data(0).n(1), 0, 0, _
                    item0(i2%).data(0).line_no(0), _
                     0, "1", "1", "1", "", "1", 0, _
                      re, 0, outI2%, 0, 0, condition_data0, False)
             If dir = "1" Then
               outp1$ = p1$
                outp2$ = p2$
              Else
               outp1$ = p2$
                outp2$ = p1$
              End If
        ElseIf item0(i1%).data(0).poi(1) = item0(i2%).data(0).poi(1) Then
         'Call set_item0(item0(i1%).data(0).poi(0), _
                item0(i1%).data(0).poi(1), item0(i2%).data(0).poi(0), _
                 item0(i2%).data(0).poi(1), "+", item0(i1%).data(0).n(0), _
                  item0(i1%).data(0).n(1), item0(i2%).data(0).n(0), _
                   item0(i2%).data(0).n(1), item0(i1%).data(0).line_no(0), _
                    item0(i1%).data(0).line_no(0), "1", "1", "1", "", "1", 0, _
                     re, 0, it0%, 0)
         If item0(i1%).data(0).n(0) > item0(i2%).data(0).n(0) Then
         tp(0) = item0(i2%).data(0).poi(0)
         tp(1) = item0(i1%).data(0).poi(0)
           Call set_item0(tp(0), tp(1), tp(0), tp(1), "*", 0, _
                  0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, _
                     re, 0, outI1%, 0, 0, condition_data0, False)
          Call set_item0(item0(i1%).data(0).poi(0), _
                item0(i1%).data(0).poi(1), tp(0), tp(1), "*", item0(i1%).data(0).n(0), _
                  item0(i1%).data(0).n(1), 0, 0, _
                    item0(i2%).data(0).line_no(0), _
                     0, "1", "1", "1", "", "1", 0, _
                      re, 0, outI2%, 0, 0, condition_data0, False)
            outp1$ = p2$
             outp2$ = time_string("2", p2$, True, False)
        dir = "-1"
         Else
         tp(0) = item0(i1%).data(0).poi(0)
         tp(1) = item0(i2%).data(0).poi(0)
         Call set_item0(tp(0), tp(1), tp(0), tp(1), "*", 0, _
                  0, 0, 0, 0, 0, "1", "1", "1", "", "1", 0, _
                     re, 0, outI1%, 0, 0, condition_data0, False)
          Call set_item0(item0(i2%).data(0).poi(0), _
                item0(i2%).data(0).poi(1), tp(0), tp(1), "*", item0(i2%).data(0).n(0), _
                  item0(i2%).data(0).n(1), 0, 0, _
                    item0(i2%).data(0).line_no(0), _
                     0, "1", "1", "1", "", "1", 0, _
                      re, 0, outI2%, 0, 0, condition_data0, False)
            outp1$ = p1$
             outp2$ = time_string("2", p1$, True, False)
         End If
        Else
         Exit Function
        End If
       'combine_two_item_for_p = set_item0(tp(0), tp(1), it0%, -1, "*", 0, 0, 0, 0, 0, _
              0, "1", "1", "1", "", "1", 0, re, 0, outI1%, 0)
        '      If dir = "1" Then
         '      outp1$ = p1$
          '    Else
           '    outp1$ = p2$
            '  End If
             '  outI2% = 0
              '  outp2$ = "0"
      ' If combine_two_item_for_p > 1 Then
       ' Exit Function
       'End If
       combine_two_item_for_p = 1
     Else
       '构成三角形
        If item0(i1%).data(0).poi(0) = item0(i2%).data(0).poi(0) Then
         tp(0) = item0(i1%).data(0).poi(0)
         tp(1) = item0(i1%).data(0).poi(1)
         tp(2) = item0(i2%).data(0).poi(1)
        ElseIf item0(i1%).data(0).poi(0) = item0(i2%).data(0).poi(1) Then
         tp(0) = item0(i1%).data(0).poi(0)
         tp(1) = item0(i1%).data(0).poi(1)
         tp(2) = item0(i2%).data(0).poi(0)
        ElseIf item0(i1%).data(0).poi(1) = item0(i2%).data(0).poi(0) Then
         tp(0) = item0(i1%).data(0).poi(1)
         tp(1) = item0(i1%).data(0).poi(0)
         tp(2) = item0(i2%).data(0).poi(1)
        ElseIf item0(i1%).data(0).poi(1) = item0(i2%).data(0).poi(1) Then
         tp(0) = item0(i1%).data(0).poi(1)
         tp(1) = item0(i1%).data(0).poi(0)
         tp(2) = item0(i2%).data(0).poi(0)
        Else
         Exit Function
        End If
       l(0) = item0(i1%).data(0).line_no(0)
       l(1) = item0(i2%).data(0).line_no(0)
       If l(0) <> l(1) Then
        If ty = 0 Then
         If is_dverti(l(0), l(1), n%, -1000, 0, 0, 0, 0) Then '直角边
          c_data.condition_no = 1
           c_data.condition(1).ty = verti_
            c_data.condition(1).no = n%
          outp1$ = p1$
           combine_two_item_for_p = set_item0(tp(1), tp(2), tp(1), tp(2), "*", 0, 0, 0, 0, 0, 0, _
              "1", "1", "1", "", "1", 0, re, 0, outI1%, 0, 0, condition_data0, False)
            If combine_two_item_for_p > 1 Then
             Exit Function
            End If
          outp2$ = "0"
           outI2% = 0
             combine_two_item_for_p = 1
         End If
        ElseIf ty = 1 Then '
         l(2) = line_number0(tp(1), tp(2), tn(1), tn(2))
         If is_dverti(l(0), l(2), n%, -1000, 0, 0, 0, 0) Then
          c_data.condition_no = 1
            c_data.condition(1).ty = verti_
             c_data.condition(1).no = n%
           outp1$ = p2$
           combine_two_item_for_p = set_item0(tp(1), tp(2), tp(1), tp(2), "*", 0, 0, 0, 0, 0, 0, _
              "1", "1", "1", "", "1", 0, re, 0, outI1%, 0, 0, condition_data0, False)
            If combine_two_item_for_p > 1 Then
             Exit Function
            End If
          outp2$ = "0"
           outI2% = 0
             combine_two_item_for_p = 1
        ElseIf is_dverti(l(1), l(2), n%, -1000, 0, 0, 0, 0) Then
           c_data.condition_no = 1
            c_data.condition(1).ty = verti_
             c_data.condition(1).no = n%
           outp1$ = p1$
           combine_two_item_for_p = set_item0(tp(1), tp(2), tp(1), tp(2), "*", 0, 0, 0, 0, 0, 0, _
              "1", "1", "1", "", "1", 0, re, 0, outI1%, 0, 0, condition_data0, False)
            If combine_two_item_for_p > 1 Then
             Exit Function
            End If
          outp2$ = "0"
           outI2% = 0
             combine_two_item_for_p = 1
        End If
       End If
      End If
     End If
   End If
 End If
End If
End Function

Public Function combine_angle_value_for_triangle(ByVal A1%, ByVal A2%, ByVal A3%) As Byte
Dim i%, n%, tn%
Dim A(2) As Integer
Dim para(2) As String
Dim tv As String
Dim A_(2) As Integer
Dim para_(2) As String
Dim tv_ As String
Dim temp_record As total_record_type
temp_record.record_data.data0.condition_data.condition_no = 1
temp_record.record_data.data0.condition_data.condition(1).ty = angle3_value_
temp_record.record_data.data0.condition_data.condition(1).no = angle(A1%).data(0).value_no
temp_record.record_data.data0.theorem_no = 1
If angle(A2%).data(0).value <> "" And angle(A2%).data(0).value <> "" Then
 Exit Function
ElseIf angle(A2%).data(0).value <> "" Then
 Call add_conditions_to_record(angle3_value_, angle(A2%).data(0).value_no, 0, 0, _
        temp_record.record_data.data0.condition_data)
 combine_angle_value_for_triangle = set_angle_value(A3%, minus_string("180", _
    add_string(angle(A1%).data(0).value, angle(A2%).data(0).value, False, False), True, False), _
       temp_record, 0, 0, False)
ElseIf angle(A3%).data(0).value <> "" Then
 Call add_conditions_to_record(angle3_value_, angle(A3%).data(0).value_no, 0, 0, _
        temp_record.record_data.data0.condition_data)
 combine_angle_value_for_triangle = set_angle_value(A2%, minus_string("180", _
    add_string(angle(A1%).data(0).value, angle(A3%).data(0).value, False, False), True, False), _
       temp_record, 0, 0, False)
Else
A(0) = A1%
A(1) = A2%
A(2) = A3%
para(0) = "1"
para(1) = "1"
para(2) = "1"
tv = minus_string("180", angle(A1%).data(0).value_no, True, False)
para(0) = "0"
For i% = 1 To 2
  n% = T_angle(angle(A(i%)).data(0).total_no).data(0).is_used_no
   A_(i%) = T_angle(angle(A(i%)).data(0).total_no).data(0).angle_no(n%).no
  If Abs(n% - angle(A(i%)).data(0).total_no_) Mod 2 = 1 Then
   para(i%) = "-1"
    tv = minus_string(tv, "180", True, False)
  End If
Next i%
If is_two_angle_value(A(1), A(2), "", "", "", "", tn%, -1000, A_(1), A_(2), _
   para(1), para(2), tv_) Then
    If A(1) = A_(1) And A_(2) = A_(2) Then
    ElseIf A(1) = A_(2) And A_(2) = A_(1) Then
     Call exchange_two_integer(A(1), A(2))
     Call exchange_string(para(1), para(2))
    Else
     Exit Function
    End If
    para(2) = time_string(para(2), para_(1), True, False)
    tv = time_string(tv, para_(1), True, False)
    para_(2) = time_string(para_(2), para(1), True, False)
    tv_ = time_string(tv_, para(1), True, False)
    tv_ = minus_string(tv_, tv, True, False)
    para_(2) = minus_string(para_(2), para(2), True, False)
    If para_(1) <> "0" Then
    combine_angle_value_for_triangle = set_angle_value(A_(2), _
      divide_string(tv_, para_(2), True, False), temp_record, 0, 0, False)
    End If
End If
End If
End Function
Public Sub read_angle_para_from_angle3_value(A3_v As angle3_value_data0_type, ByVal m%, A1%, A2%, A3%, _
         p1 As String, p2 As String, p3 As String, v As String, ty As Byte)
Dim A(3) As Integer 'ty=0 A1,A3 ty=1 a2 a3
Dim pA(2) As String
Dim tv As String
 If m% < 3 Then
   A1% = A3_v.angle(m%)
   A2% = A3_v.angle((m% + 1) Mod 3)
   A3% = A3_v.angle((m% + 2) Mod 3)
   p1 = A3_v.para(m%)
   p2 = A3_v.para((m% + 1) Mod 3)
   p3 = A3_v.para((m% + 2) Mod 3)
   v = A3_v.value
ElseIf m% = 3 Then
 If ty = 0 Then
 If A3_v.ty(0) = 3 Or A3_v.ty(0) = 5 Then 'A+B-C=0
  A1% = A3_v.angle(3)
  A2% = A3_v.angle(0)
  A3% = A3_v.angle(2)
  p2 = minus_string(A3_v.para(0), A3_v.para(1), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = A3_v.para(1)
  p3 = A3_v.para(2)
  v = A3_v.value
 ElseIf A3_v.ty(0) = 4 Or A3_v.ty(0) = 8 Then 'A-B=C
  A2% = A3_v.angle(0)
  A1% = A3_v.angle(3)
  A3% = A3_v.angle(2)
  p2 = add_string(A3_v.para(0), A3_v.para(1), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = time_string("-1", A3_v.para(1), True, False)
  p3 = A3_v.para(2)
  v = A3_v.value
 ElseIf A3_v.ty(0) = 6 Or A3_v.ty(0) = 7 Then 'A-B+C=0
  A2% = A3_v.angle(0)
  A1% = A3_v.angle(3)
  A3% = A3_v.angle(2)
  p2 = add_string(A3_v.para(0), A3_v.para(1), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = A3_v.para(1)
  p3 = A3_v.para(2)
  v = A3_v.value
 ElseIf A3_v.ty(0) = 15 Or A3_v.ty(0) = 17 Then 'A+B+C=180
  A2% = A3_v.angle(0)
  A1% = A3_v.angle(3)
  A3% = A3_v.angle(2)
  v = minus_string(A3_v.value, time_string("180", A3_v.para(1), False, False), True, False)
  p2 = minus_string(A3_v.para(0), A3_v.para(1), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = time_string(A3_v.para(1), "-1", True, False)
  p3 = A3_v.para(2)
 ElseIf A3_v.ty(0) = 16 Or A3_v.ty(0) = 18 Then 'A+B-C=180
  A2% = A3_v.angle(0)
  A1% = A3_v.angle(3)
  A3% = A3_v.angle(2)
  v = minus_string(A3_v.value, time_string("180", A3_v.para(1), False, False), True, False)
  p2 = minus_string(A3_v.para(0), A3_v.para(1), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = A3_v.para(1)
  p3 = A3_v.para(2)
 ElseIf A3_v.ty(0) = 23 Then 'A-B+C=180
  A2% = A3_v.angle(0)
  A1% = A3_v.angle(3)
  A3% = A3_v.angle(2)
  v = add_string(A3_v.value, time_string("180", A3_v.para(1), False, False), True, False)
  p2 = add_string(A3_v.para(0), A3_v.para(1), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = A3_v.para(1)
  p3 = A3_v.para(2)
 ElseIf A3_v.ty(0) = 24 Then 'A-B-C=-180
  A2% = A3_v.angle(0)
  A1% = A3_v.angle(3)
  A3% = A3_v.angle(2)
  v = minus_string(A3_v.value, time_string("180", A3_v.para(1), False, False), True, False)
  p2 = add_string(A3_v.para(0), A3_v.para(1), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = time_string("-1", A3_v.para(1), True, False)
  p3 = A3_v.para(2)
 ElseIf A3_v.ty(0) = 9 Or A3_v.ty(0) = 10 Then 'A+B+C=360
  A2% = A3_v.angle(0)
  A1% = A3_v.angle(3)
  A3% = A3_v.angle(2)
  v = minus_string(A3_v.value, time_string("360", A3_v.para(1), False, False), True, False)
  p2 = minus_string(A3_v.para(0), A3_v.para(1), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = time_string(A3_v.para(1), "-1", True, False)
  p3 = A3_v.para(2)
 End If
 Else 'ty=1
  If A3_v.ty(0) = 3 Or A3_v.ty(0) = 5 Then 'A+B=C
  A1% = A3_v.angle(3)
  A2% = A3_v.angle(1)
  A3% = A3_v.angle(2)
  p2 = minus_string(A3_v.para(1), A3_v.para(0), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = A3_v.para(0)
  p3 = A3_v.para(2)
  v = A3_v.value
 ElseIf A3_v.ty(0) = 4 Or A3_v.ty(0) = 8 Then 'A-B=C
  A2% = A3_v.angle(1)
  A1% = A3_v.angle(3)
  A3% = A3_v.angle(2)
  p2 = add_string(A3_v.para(1), A3_v.para(0), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = A3_v.para(0)
  p3 = A3_v.para(2)
  v = A3_v.value
 ElseIf A3_v.ty(0) = 6 Or A3_v.ty(0) = 7 Then 'A-B+C=0
  A2% = A3_v.angle(1)
  A1% = A3_v.angle(3)
  A3% = A3_v.angle(2)
  p2 = add_string(A3_v.para(1), A3_v.para(0), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = time_string(A3_v.para(0), "-1", True, False)
  p3 = A3_v.para(2)
  v = A3_v.value
 ElseIf A3_v.ty(0) = 15 Or A3_v.ty(0) = 17 Then 'A+B+C=180
  A2% = A3_v.angle(1)
  A1% = A3_v.angle(3)
  A3% = A3_v.angle(2)
  v = minus_string(A3_v.value, time_string("180", A3_v.para(0), False, False), True, False)
  p2 = minus_string(A3_v.para(1), A3_v.para(0), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = time_string(A3_v.para(0), "-1", True, False)
  p3 = A3_v.para(2)
 ElseIf A3_v.ty(0) = 16 Or A3_v.ty(0) = 18 Then 'A+B-C=180
  A2% = A3_v.angle(1)
  A1% = A3_v.angle(3)
  A3% = A3_v.angle(2)
  v = minus_string(A3_v.value, time_string("180", A3_v.para(0), False, False), True, False)
  p2 = minus_string(A3_v.para(1), A3_v.para(0), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = A3_v.para(0)
  p3 = A3_v.para(2)
 ElseIf A3_v.ty(0) = 23 Then 'A-B+C=180
  A2% = A3_v.angle(1)
  A1% = A3_v.angle(3)
  A3% = A3_v.angle(2)
  v = minus_string(A3_v.value, time_string("180", A3_v.para(0), False, False), True, False)
  p2 = add_string(A3_v.para(0), A3_v.para(1), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = time_string(A3_v.para(0), "-1", True, False)
  p3 = A3_v.para(2)
 ElseIf A3_v.ty(0) = 24 Then 'A-B-C=-180
  A2% = A3_v.angle(1)
  A1% = A3_v.angle(3)
  A3% = A3_v.angle(2)
  v = add_string(A3_v.value, time_string("180", A3_v.para(0), False, False), True, False)
  p2 = add_string(A3_v.para(0), A3_v.para(1), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = A3_v.para(0)
  p3 = A3_v.para(2)
 ElseIf A3_v.ty(0) = 9 Or A3_v.ty(0) = 10 Then 'A+B+C=360
  A2% = A3_v.angle(1)
  A1% = A3_v.angle(3)
  A3% = A3_v.angle(2)
  v = minus_string(A3_v.value, time_string("360", A3_v.para(0), False, False), True, False)
  p2 = minus_string(A3_v.para(1), A3_v.para(0), True, False)
  If p2 = "0" Then
  A2% = 0
  End If
  p1 = time_string(A3_v.para(0), "-1", True, False)
  p3 = A3_v.para(2)
 End If
 If A3_v.angle_(3) <> A1% Then
    A1% = A3_v.angle_(3)
 End If
End If
End If
End Sub

Public Function combine_mid_point_with_general_string(ByVal p1%, ByVal p2%, ByVal p3%, _
  gs%, re As total_record_type) As Byte
Dim i%, n%
Dim p(1) As Integer
Dim ty As Byte
Dim tl(1) As Integer
Dim tn(3) As Integer
Dim tp(3) As Integer
Dim it(2) As Integer
Dim pA(2) As String
Dim tp_(2) As Integer
Dim sig As String
Dim midpoint_data0 As mid_point_data0_type
Dim temp_record As total_record_type
If gs% = 0 Then
For i% = 1 To last_conditions.last_cond(1).general_string_no
 ty = is_general_string_satis_midpoint(i%, tp(0), tp(1), tp(2), tp(3), p(0), p(1), pA(2), _
       it(0), it(1), pA(0), pA(1))
 If ty > 0 Then
  combine_mid_point_with_general_string = combine_mid_point_with_general_string0( _
      tp(0), tp(1), tp(2), tp(3), p(0), p(1), pA(2), p1%, p2%, p3%, it(0), it(1), pA(0), _
       pA(1), i%, re, ty)
 If combine_mid_point_with_general_string > 1 Then
  Exit Function
 End If
 End If
Next i%
ElseIf p1% = 0 And p2% = 0 And p3% = 0 Then
 ty = is_general_string_satis_midpoint(gs%, tp(0), tp(1), tp(2), tp(3), p(0), p(1), pA(2), _
       it(0), it(1), pA(0), pA(1))
 If ty > 0 Then
  tl(0) = line_number0(tp(0), tp(1), tn(0), tn(1))
  If tn(0) > tn(1) Then
  Call exchange_two_integer(tn(0), tn(1))
  Call exchange_two_integer(tp(0), tp(1))
  End If
  tl(1) = line_number0(tp(2), tp(3), tn(2), tn(3))
  If tn(0) > tn(1) Then
  Call exchange_two_integer(tn(2), tn(3))
  Call exchange_two_integer(tp(2), tp(3))
  End If
If ty = 1 Then
  If tp(0) = tp(2) Then
    Call exchange_two_integer(tp(0), tp(1))
    Call exchange_two_integer(tp(2), tp(3))
  ElseIf tp(0) = tp(3) Then
    Call exchange_two_integer(tp(0), tp(1))
  ElseIf tp(1) = tp(2) Then
   Call exchange_two_integer(tp(2), tp(3))
  End If
  If is_mid_point(tp(0), 0, tp(2), 0, 0, 0, 0, n%, -1000, 0, 0, 0, 0, 0, 0, _
       midpoint_data0, "", 0, 0, 0, temp_record.record_data.data0.condition_data) Or tp(0) = tp(2) Then
        If tp(0) = tp(2) Then
         tp_(0) = tp(0)
         tp_(1) = tp(0)
         tp_(2) = tp(2)
        Else
         tp_(0) = midpoint_data0.poi(0)
         tp_(1) = midpoint_data0.poi(1)
         tp_(2) = midpoint_data0.poi(2)
        End If
  combine_mid_point_with_general_string = combine_mid_point_with_general_string0( _
      tp(0), tp(1), tp(2), tp(3), p(0), p(1), pA(2), tp_(0), tp_(1), tp_(2), _
          it(0), it(1), pA(0), pA(1), gs%, temp_record, ty)
    If combine_mid_point_with_general_string > 1 Then
     Exit Function
    End If
  ElseIf is_mid_point(tp(1), 0, tp(3), 0, 0, 0, 0, n%, -1000, 0, 0, 0, 0, 0, 0, _
       midpoint_data0, "", 0, 0, 0, temp_record.record_data.data0.condition_data) Or tp(1) = tp(3) Then
        If tp(1) = tp(3) Then
         tp_(0) = tp(1)
         tp_(1) = tp(1)
         tp_(2) = tp(3)
        Else
         tp_(0) = midpoint_data0.poi(0)
         tp_(1) = midpoint_data0.poi(1)
         tp_(2) = midpoint_data0.poi(2)
        End If
  combine_mid_point_with_general_string = combine_mid_point_with_general_string0( _
      tp(0), tp(1), tp(2), tp(3), p(0), p(1), pA(2), tp_(0), tp_(1), tp_(2), _
          it(0), it(1), pA(0), pA(1), gs%, temp_record, ty)
    If combine_mid_point_with_general_string > 1 Then
      Exit Function
    End If
  End If
 Else 'ty=2
  If is_mid_point(tp(0), tp(2), 0, 0, 0, 0, 0, n%, -1000, 0, 0, 0, 0, 0, 0, _
       midpoint_data0, "", 0, 0, 0, temp_record.record_data.data0.condition_data) Or tp(0) = tp(2) Then
        If tp(0) = tp(2) Then
         tp_(0) = tp(0)
         tp_(1) = tp(0)
         tp_(2) = tp(2)
        Else
         tp_(0) = midpoint_data0.poi(0)
         tp_(1) = midpoint_data0.poi(1)
         tp_(2) = midpoint_data0.poi(2)
        End If
  combine_mid_point_with_general_string = combine_mid_point_with_general_string0( _
      tp(0), tp(1), tp(2), tp(3), p(0), p(1), pA(2), tp_(0), tp_(1), tp_(2), _
          it(0), it(1), pA(0), pA(1), gs%, temp_record, ty)
    If combine_mid_point_with_general_string > 1 Then
     Exit Function
    End If
 ElseIf is_mid_point(tp(1), tp(3), 0, 0, 0, 0, 0, n%, -1000, 0, 0, 0, 0, 0, 0, _
       midpoint_data0, "", 0, 0, 0, temp_record.record_data.data0.condition_data) Or tp(1) = tp(3) Then
        If tp(1) = tp(3) Then
         tp_(0) = tp(1)
         tp_(1) = tp(1)
         tp_(2) = tp(3)
        Else
         tp_(0) = midpoint_data0.poi(0)
         tp_(1) = midpoint_data0.poi(1)
         tp_(2) = midpoint_data0.poi(2)
        End If
  combine_mid_point_with_general_string = combine_mid_point_with_general_string0( _
      tp(0), tp(1), tp(2), tp(3), p(0), p(1), pA(2), tp_(0), tp_(1), tp_(2), _
          it(0), it(1), pA(0), pA(1), gs%, temp_record, ty)
    If combine_mid_point_with_general_string > 1 Then
      Exit Function
    End If
 End If
 End If
 End If
End If
End Function

Public Function combine_mid_point_with_general_string0(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
                ByVal p_1%, ByVal p_2%, ByVal pa3$, ByVal mp1%, ByVal mp2%, ByVal mp3%, ByVal it1%, _
                 ByVal it2%, ByVal pA1$, ByVal pA2$, gs%, re As total_record_type, operat_type As Byte) As Byte
                 'operat_ty=1,liang bian , =2bian+minline
Dim n(2) As Integer
Dim tl(1) As Integer
Dim tn(4) As Integer
Dim mp_%, it3%
Dim temp_record As total_record_type
Dim mid_point_data0 As mid_point_data0_type
 tl(0) = line_number0(p1%, p2%, tn(0), tn(1))
  If tn(0) > tn(1) Then
   Call exchange_two_integer(p1%, p2%)
   Call exchange_two_integer(tn(0), tn(1))
  End If
 tl(1) = line_number0(p3%, p4%, tn(2), tn(3))
  If tn(2) > tn(3) Then
   Call exchange_two_integer(p3%, p4%)
   Call exchange_two_integer(tn(2), tn(3))
  End If
   If is_dparal(tl(0), tl(1), n(0), -1000, 0, 0, 0, 0) Or tl(0) = tl(1) Then '平行 '梯形
    If operat_type = 1 Then
      If is_same_two_point(p1%, p3%, mp1%, mp3%) Then '中点
       If is_mid_point(p2%, mp_%, p4%, 0, 0, 0, 0, n(1), -1000, 0, 0, 0, 0, 0, 0, _
           mid_point_data0, "", 0, 0, 0, condition_data0) Or p2% = p4% Then '另一中点
            If p2% = p4% Then
             mp_% = p2%
            End If
       Else
       If last_conditions.last_cond(1).new_midpoint_no = 0 Then
          last_conditions.last_cond(1).new_midpoint_no = 1
           Call set_pseudo_mid_point(p2%, 0, p4%)
            Exit Function
        '    combine_mid_point_with_general_string0 = add_mid_point(p2%, 0, p4%, 0)
        '   If combine_mid_point_with_general_string0 > 1 Then
        '    Exit Function
        ' Else
        '  GoTo combine_mid_point_with_general_string0_mark0
        ' End If
        Else
         GoTo combine_mid_point_with_general_string0_mark0
        End If
       End If
      ElseIf is_same_two_point(p2%, p4%, mp1%, mp3%) Then
       If is_mid_point(p1%, mp_%, p3%, 0, 0, 0, 0, n(1), -1000, 0, 0, 0, 0, 0, 0, _
           mid_point_data0, "", 0, 0, 0, condition_data0) Or p1% = p3% Then
            If p1% = p3% Then
             mp_% = p1%
            End If
       Else
        If last_conditions.last_cond(1).new_midpoint_no = 0 Then
          last_conditions.last_cond(1).new_midpoint_no = 1
           Call set_pseudo_mid_point(p1%, 0, p3%)
            Exit Function
         'combine_mid_point_with_general_string0 = add_mid_point(p1%, 0, p3%, 0)
         'If combine_mid_point_with_general_string0 > 1 Then
         ' Exit Function
         'Else
         ' GoTo combine_mid_point_with_general_string0_mark0
         'End If
        Else
         GoTo combine_mid_point_with_general_string0_mark0
        End If
       End If
      Else
       GoTo combine_mid_point_with_general_string0_mark0
      End If
    temp_record = re
    Call add_conditions_to_record(paral_, n(0), 0, 0, temp_record.record_data.data0.condition_data)
    Call add_conditions_to_record(midpoint_, n(1), 0, 0, temp_record.record_data.data0.condition_data)
    If tl(0) <> tl(1) Then
    temp_record.record_data.data0.theorem_no = 98
    End If
    n(2) = 0
    combine_mid_point_with_general_string0 = set_three_line_value(p1%, p2%, p3%, p4%, mp_%, mp2%, _
        0, 0, 0, 0, 0, 0, 0, 0, 0, "1", "1", "-2", "0", temp_record, n(2), 0, 0)
    If item0(it1%).data(0).sig = "~" Then
    combine_mid_point_with_general_string0 = set_item0(mp_%, mp2%, 0, 0, "~", 0, 0, 0, 0, 0, 0, _
         "1", "1", "1", "", "1", _
             0, condition_data0, 0, it3%, 0, 0, condition_data0, False)
      If combine_mid_point_with_general_string0 > 1 Then
       Exit Function
      End If
    Else
    combine_mid_point_with_general_string0 = set_item0(mp_%, mp2%, p_1%, p_2%, "/", _
      0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", _
             0, condition_data0, 0, it3, 0, 0, condition_data0, False)
      If combine_mid_point_with_general_string0 > 1 Then
       Exit Function
      End If
    End If
   Else 'bian+ midline
   '*****
      If (p1% = mp2% And (p3% = mp1% Or p3% = mp3%)) Or p1% = p3% Then
        If p1% = p2% Then
           mp3% = p1%
        'ElseIf p3% = mp1% Then
        ElseIf p3% = mp3% Then
         mp3% = mp1%
        End If
        If is_mid_point(p4%, p2%, mp_%, 0, 0, 0, 0, n(1), -1000, 0, 0, 0, 0, 0, 0, _
           mid_point_data0, "", 0, 0, 0, condition_data0) Or p2% = p4% Then
            If p2% = p4% Then
             mp_% = p2%
            ElseIf mid_point_data0.poi(0) = p4% Then
             mp_% = mid_point_data0.poi(2)
            Else
             mp_% = mid_point_data0.poi(0)
            End If
       Else
       If last_conditions.last_cond(1).new_midpoint_no = 0 Then
          last_conditions.last_cond(1).new_midpoint_no = 1
          Call set_pseudo_mid_point(p4%, 0, p2%)
           Exit Function
        'combine_mid_point_with_general_string0 = add_mid_point(p4%, p2%, 0,0)
        ' If combine_mid_point_with_general_string0 > 1 Then
        '  Exit Function
        ' Else
        '  GoTo combine_mid_point_with_general_string0_mark0
        ' End If
        Else
         GoTo combine_mid_point_with_general_string0_mark0
        End If
       End If
      ElseIf (p2% = mp3% And (p4% = mp1% Or p4% = mp3%)) Or p2% = p4% Then
       If p2% = p4% Then
          mp3% = p2%
       ElseIf p4% = mp3% Then
          mp3% = mp1%
       End If
       If is_mid_point(p3%, p1%, mp_%, 0, 0, 0, 0, n(1), -1000, 0, 0, 0, 0, 0, 0, _
           mid_point_data0, "", 0, 0, 0, condition_data0) Or p1% = p3% Then
            If p1% = p3% Then
             mp_% = p1%
            ElseIf mid_point_data0.poi(0) = p4% Then
             mp_% = mid_point_data0.poi(2)
            Else
             mp_% = mid_point_data0.poi(0)
            End If
       Else
        If last_conditions.last_cond(1).new_midpoint_no = 0 Then
          last_conditions.last_cond(1).new_midpoint_no = 1
          Call set_pseudo_mid_point(p3%, p1%, 0)
           Exit Function
         'combine_mid_point_with_general_string0 = add_mid_point(p3%, p1%, 0, 0)
         'If combine_mid_point_with_general_string0 > 1 Then
         ' Exit Function
         'Else
         ' GoTo combine_mid_point_with_general_string0_mark0
         'End If
        Else
         GoTo combine_mid_point_with_general_string0_mark0
        End If
       End If
      Else
       GoTo combine_mid_point_with_general_string0_mark0
      End If
    temp_record = re
    Call add_conditions_to_record(paral_, n(0), 0, 0, temp_record.record_data.data0.condition_data)
    Call add_conditions_to_record(midpoint_, n(1), 0, 0, temp_record.record_data.data0.condition_data)
    If tl(0) <> tl(1) Then
    temp_record.record_data.data0.theorem_no = 98
    Else
    temp_record.record_data.data0.theorem_no = 1
    End If
    n(2) = 0
    combine_mid_point_with_general_string0 = set_three_line_value(p1%, p2%, p3%, p4%, mp_%, mp3%, _
        0, 0, 0, 0, 0, 0, 0, 0, 0, "2", "-1", "-1", "0", temp_record, n(2), 0, 0)
    If item0(it1%).data(0).sig = "~" Then
    Call set_item0(mp_%, mp3%, 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", _
             0, condition_data0, 0, it3%, 0, 0, condition_data0, False)
    Else
    Call set_item0(mp_%, mp3%, p_1%, p_2%, "/", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", "1", _
             0, condition_data0, 0, it3, 0, 0, condition_data0, False)
    End If
      If combine_mid_point_with_general_string0 > 1 Then
       Exit Function
      End If
   '*****
   End If
    'temp_record = re
    Call add_conditions_to_record(general_string_, gs%, 0, 0, temp_record.record_data.data0.condition_data)
    'temp_record.record_data.data0.condition_data.condition_no = 2
    'temp_record.record_data.data0.condition_data.condition(1).ty = line3_value_
    'temp_record.record_data.data0.condition_data.condition(1).no = n(2)
    'temp_record.record_data.data0.condition_data.condition(2).ty = general_string_
    'temp_record.record_data.data0.condition_data.condition(2).no = gs%
    temp_record.record_data.data0.theorem_no = 1
   combine_mid_point_with_general_string0 = set_general_string(it3, it1%, it2%, 0, pa3, pA1, pA2, _
      "0", "", general_string(gs%).record_.conclusion_no, 0, 0, temp_record, 0, 0)
    If combine_mid_point_with_general_string0 > 1 Then
     Exit Function
    End If
  End If
combine_mid_point_with_general_string0_mark0:
End Function
Public Function sub_line_value_to_line_value(ByVal p1%, ByVal p2%, _
     ByVal n1%, ByVal n2%, ByVal l1%, ByVal para$, ByVal v1$, ByVal p3%, ByVal p4%, _
       ByVal n3%, ByVal n4%, ByVal l2%, ByVal v2$, o_p1%, o_p2%, o_n1%, o_n2%, o_l%, o_para$, o_V$) As Boolean
Dim ty As Integer
o_p1% = p1%
o_p2% = p2%
o_n1% = n1%
o_n2% = n2%
o_l% = l1%
o_para$ = para$
o_V$ = v1$
If l1% <> l2% Then
 Exit Function
ElseIf is_contain_x(v2$, "x", 1) = False Then
 Exit Function
Else
 If n1% = n3% And n2% = n4% Then
   o_p1% = 0
   o_p2% = 0
   o_n1% = 0
   o_n2% = 0
   o_l% = 0
   o_V$ = minus_string(v1$, v2$, True, False)
   sub_line_value_to_line_value = True
Else
 Call different_two_line(p3%, p4%, n3%, n4%, p1%, p2%, n1%, n2%, o_p1%, o_p2%, o_n1%, o_n2%, ty)
 If ty = 0 Then
  o_p1% = p1%
  o_p2% = p2%
  o_n1% = n1%
  o_n2% = n2%
  o_l% = l1%
  o_para$ = para$
  o_V$ = v1$
ElseIf ty = -1 Then
   o_para = para$
   o_V$ = minus_string(v1$, v2$, True, False)
   sub_line_value_to_line_value = True
ElseIf ty = 1 Then
   o_para = time_string("-1", para$, True, False)
   o_V$ = minus_string(v1$, v2$, True, False)
   sub_line_value_to_line_value = True
End If
End If
End If
End Function

Public Function subs_line_value_to_two_line_value(ByVal l2_v%, ByVal lv%) As Byte
Dim t_l As two_line_value_data0_type
Dim temp_record As total_record_type
  If two_line_value(l2_v%).data(0).data0.para(0) = "1" And two_line_value(l2_v%).data(0).data0.para(1) = "1" Then
   t_l = two_line_value(l2_v%).data(0).data0
   If sub_line_value_to_line_value(t_l.poi(0), t_l.poi(1), t_l.n(0), t_l.n(1), t_l.line_no(0), _
       t_l.para(0), t_l.value, line_value(lv%).data(0).data0.poi(0), line_value(lv%).data(0).data0.poi(1), _
        line_value(lv%).data(0).data0.n(0), line_value(lv%).data(0).data0.n(1), _
         line_value(lv%).data(0).data0.line_no, line_value(lv%).data(0).data0.value_, _
          t_l.poi(0), t_l.poi(1), t_l.n(0), t_l.n(1), t_l.line_no(0), t_l.para(0), t_l.value_) Then
     temp_record.record_data.data0.condition_data.condition_no = 2
     temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(1).no = lv%
     temp_record.record_data.data0.condition_data.condition(1).ty = two_line_value_
     temp_record.record_data.data0.condition_data.condition(1).no% = l2_v%
   subs_line_value_to_two_line_value = set_two_line_value(t_l.poi(0), t_l.poi(1), _
     t_l.poi(2), t_l.poi(3), t_l.n(0), t_l.n(1), t_l.n(2), _
      t_l.n(3), t_l.line_no(0), t_l.line_no(1), t_l.para(0), _
       t_l.para(1), t_l.value_, temp_record, 0, 0)
    If subs_line_value_to_two_line_value > 1 Then
     Exit Function
    End If
    ElseIf sub_line_value_to_line_value(t_l.poi(2), t_l.poi(3), t_l.n(2), t_l.n(3), t_l.line_no(1), _
       t_l.para(1), t_l.value_, line_value(lv%).data(0).data0.poi(0), line_value(lv%).data(0).data0.poi(1), _
        line_value(lv%).data(0).data0.n(0), line_value(lv%).data(0).data0.n(1), _
         line_value(lv%).data(0).data0.line_no, line_value(lv%).data(0).data0.value, _
          t_l.poi(2), t_l.poi(3), t_l.n(2), t_l.n(3), t_l.line_no(1), t_l.para(1), t_l.value_) Then
     temp_record.record_data.data0.condition_data.condition_no = 2
     temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(1).no = lv%
     temp_record.record_data.data0.condition_data.condition(1).ty = two_line_value_
     temp_record.record_data.data0.condition_data.condition(1).no% = l2_v%
   subs_line_value_to_two_line_value = set_two_line_value(t_l.poi(0), t_l.poi(1), _
     t_l.poi(2), t_l.poi(3), t_l.n(0), t_l.n(1), t_l.n(2), _
      t_l.n(3), t_l.line_no(0), t_l.line_no(1), t_l.para(0), _
       t_l.para(1), t_l.value_, temp_record, 0, 0)
    If subs_line_value_to_two_line_value > 1 Then
     Exit Function
    End If
   End If
  End If

End Function

Public Function subs_line_value_to_line3_value(ByVal l3_v%, ByVal lv%) As Byte
Dim t_l As line3_value_data0_type
Dim temp_record As total_record_type
  If line3_value(l3_v%).data(0).data0.para(0) = "1" And line3_value(l3_v%).data(0).data0.para(1) = "1" And _
        line3_value(l3_v%).data(0).data0.para(2) = "1" Then
   t_l = line3_value(l3_v%).data(0).data0
   If sub_line_value_to_line_value(t_l.poi(0), t_l.poi(1), t_l.n(0), t_l.n(1), t_l.line_no(0), _
       t_l.para(0), t_l.value_, line_value(lv%).data(0).data0.poi(0), line_value(lv%).data(0).data0.poi(1), _
        line_value(lv%).data(0).data0.n(0), line_value(lv%).data(0).data0.n(1), _
         line_value(lv%).data(0).data0.line_no, line_value(lv%).data(0).data0.value_, _
          t_l.poi(0), t_l.poi(1), t_l.n(0), t_l.n(1), t_l.line_no(0), t_l.para(0), t_l.value_) Or _
     sub_line_value_to_line_value(t_l.poi(2), t_l.poi(3), t_l.n(2), t_l.n(3), t_l.line_no(1), _
       t_l.para(1), t_l.value_, line_value(lv%).data(0).data0.poi(0), line_value(lv%).data(0).data0.poi(1), _
        line_value(lv%).data(0).data0.n(0), line_value(lv%).data(0).data0.n(1), _
         line_value(lv%).data(0).data0.line_no, line_value(lv%).data(0).data0.value_, _
          t_l.poi(2), t_l.poi(3), t_l.n(2), t_l.n(3), t_l.line_no(1), t_l.para(1), t_l.value_) Or _
     sub_line_value_to_line_value(t_l.poi(4), t_l.poi(5), t_l.n(4), t_l.n(5), t_l.line_no(2), _
       t_l.para(2), t_l.value_, line_value(lv%).data(0).data0.poi(0), line_value(lv%).data(0).data0.poi(1), _
        line_value(lv%).data(0).data0.n(0), line_value(lv%).data(0).data0.n(1), _
         line_value(lv%).data(0).data0.line_no, line_value(lv%).data(0).data0.value_, _
          t_l.poi(4), t_l.poi(5), t_l.n(4), t_l.n(5), t_l.line_no(2), t_l.para(2), t_l.value_) Then
     temp_record.record_data.data0.condition_data.condition_no = 2
     temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
     temp_record.record_data.data0.condition_data.condition(1).no = lv%
     temp_record.record_data.data0.condition_data.condition(2).ty = line3_value_
     temp_record.record_data.data0.condition_data.condition(2).no% = l3_v%
   subs_line_value_to_line3_value = set_three_line_value(t_l.poi(0), t_l.poi(1), _
     t_l.poi(2), t_l.poi(3), t_l.poi(4), t_l.poi(5), t_l.n(0), t_l.n(1), t_l.n(2), _
      t_l.n(3), t_l.n(4), t_l.n(5), t_l.line_no(0), t_l.line_no(1), t_l.line_no(2), t_l.para(0), _
       t_l.para(1), t_l.para(2), t_l.value_, temp_record, 0, 0, 0)
    If subs_line_value_to_line3_value > 1 Then
     Exit Function
    End If
   End If
  End If

End Function

Public Function combine_eangle_for_tri_function(ByVal no%) As Byte
Dim i%, n%
Dim tA(1) As Integer
Dim temp_record(1) As total_record_type
Dim tri_f As tri_function_data_type
tri_f = tri_function(no%).data(0)
temp_record(0).record_data.data0.theorem_no = 1
temp_record(0).record_data.data0.condition_data.condition_no = 1
temp_record(0).record_data.data0.condition_data.condition(1).ty = tri_function_
temp_record(0).record_data.data0.condition_data.condition(1).no = no%
If angle(tri_function(no%).data(0).A).data(0).value_no > 0 Then
 For i% = 1 To last_conditions.last_cond(1).angle_value_no
  n% = angle_value.av_no(i%).no
   If angle(angle3_value(n%).data(0).data0.angle(0)).data(0).value = _
       angle(tri_f.A).data(0).value Then
        temp_record(1) = temp_record(0)
         Call add_conditions_to_record(angle3_value_, _
            angle(tri_f.A).data(0).value_no, n%, 0, _
              temp_record(1).record_data.data0.condition_data)
      combine_eangle_for_tri_function = set_tri_function( _
         angle3_value(n%).data(0).data0.angle(0), "", "", "", "", _
          0, temp_record(1), False, tri_f, 0)
       If combine_eangle_for_tri_function > 1 Then
         Exit Function
       End If
   End If
 Next i%
Else
 For i% = 1 To last_conditions.last_cond(1).eangle_no
   temp_record(1) = temp_record(0)
    n% = Deangle.av_no(i%).no
     tA(0) = angle3_value(n%).data(0).data0.angle(0)
      tA(1) = angle3_value(n%).data(0).data0.angle(1)
  If tA(0) = tri_f.A Then
         Call add_conditions_to_record(angle3_value_, _
             n%, 0, 0, _
              temp_record(1).record_data.data0.condition_data)
      combine_eangle_for_tri_function = set_tri_function( _
         tA(1), "", "", "", "", _
          0, temp_record(1), False, tri_f, 0)
       If combine_eangle_for_tri_function > 1 Then
         Exit Function
       End If
   ElseIf tA(1) = tri_f.A Then
         Call add_conditions_to_record(angle3_value_, _
            n%, 0, 0, _
              temp_record(1).record_data.data0.condition_data)
      combine_eangle_for_tri_function = set_tri_function( _
         tA(0), "", "", "", "", _
          0, temp_record(1), False, tri_f, 0)
       If combine_eangle_for_tri_function > 1 Then
         Exit Function
       End If
  End If
 Next i%
End If
End Function

Public Function combine_general_string_with_general_string0(ByVal it1%, ByVal it2%, ByVal it3%, ByVal it4%, _
                    g_s2 As general_string_data_type, re As condition_data_type, out_gs As general_string_data_type) As Boolean
Dim t_g_s(1) As general_string_data_type
Dim i%, j%, n%, m%, no%
Dim t_s(1) As String
out_gs = g_s2
t_g_s(0).item(0) = it1%
t_g_s(0).item(1) = it2%
t_g_s(0).item(2) = it3%
t_g_s(0).item(3) = it4%
   If search_for_general_string(t_g_s(0), 0, no%, 0) Then
    If general_string(no%).data(0).value <> "" Then
        t_g_s(0) = general_string(no%).data(0)
 t_g_s(1) = g_s2
 n% = -1
 m% = -1
 For i% = 3 To 0 Step -1
 For j% = 3 To 0 Step -1
  If t_g_s(0).item(i%) = t_g_s(1).item(j%) And t_g_s(0).item(i%) > 0 Then
   n% = i%
   m% = j%
   GoTo combine_general_string_general_string0_mark0
  End If
 Next j%
Next i%
combine_general_string_general_string0_mark0:
t_s(0) = t_g_s(0).para(n%)
t_s(1) = t_g_s(1).para(m%)
If t_s(0) <> t_s(1) Then
If t_g_s(0).value = "" Then
 For i% = 0 To 3
  t_g_s(1).para(i%) = divide_string(t_g_s(1).para(i%), t_s(1), True, False)
  t_g_s(1).para(i%) = time_string(t_g_s(1).para(i%), t_s(0), True, False)
 Next i%
  t_g_s(1).value = divide_string(t_g_s(1).value, t_s(1), True, False)
  t_g_s(1).value = time_string(t_g_s(1).value, t_s(0), True, False)
Else
 For i% = 0 To 3
  t_g_s(0).para(i%) = time_string(t_g_s(0).para(i%), t_s(1), True, False)
  t_g_s(1).para(i%) = time_string(t_g_s(1).para(i%), t_s(0), True, False)
 Next i%
  t_g_s(0).value = time_string(t_g_s(0).value, t_s(1), True, False)
  t_g_s(1).value = time_string(t_g_s(1).value, t_s(0), True, False)
End If
End If
For i% = 0 To 3
For j% = 0 To 3
 If t_g_s(0).item(i%) = t_g_s(1).item(j%) And _
        (t_g_s(1).para(j%) <> "0" And t_g_s(1).para(j%) <> "") Then
    t_g_s(1).para(j%) = minus_string(t_g_s(1).para(j%), t_g_s(0).para(i%), True, False)
    t_g_s(0).item(i%) = 0
    t_g_s(0).para(i%) = "0"
    If t_g_s(1).para(j%) = "0" Then
       t_g_s(1).item(j%) = 0
    End If
 End If
Next j%
Next i%
If t_g_s(0).value = "0" Then
 Call add_conditions_to_record(general_string_, no%, 0, 0, re)
 combine_general_string_with_general_string0 = True
  out_gs = t_g_s(1)
  Exit Function
Else
   If t_g_s(0).value <> "" Then
    t_g_s(1).value = minus_string(t_g_s(1).value, t_g_s(0).value, True, False)
     Call add_conditions_to_record(general_string_, no%, 0, 0, re)
      combine_general_string_with_general_string0 = True
       out_gs = t_g_s(1)
       Exit Function
   Else
    For i% = 0 To 3
     If t_g_s(0).item(0) = 0 Then
      t_g_s(0).para(i%) = add_string(t_g_s(1).para(i%), t_g_s(0).value, True, False)
       Call add_conditions_to_record(general_string_, no%, 0, 0, re)
        combine_general_string_with_general_string0 = True
          out_gs = t_g_s(1)
      Exit Function
     End If
    Next i%
     combine_general_string_with_general_string0 = True
      out_gs = t_g_s(1)
      Exit Function
   End If
End If
End If
End If
End Function
 
Public Function combine_item0_with_two_line_value(ByVal no%) As Byte
Dim i%, j%, k%, n_%
Dim tn(1) As Integer
Dim n() As Integer
Dim last_n%
Dim t_l_v As two_line_value_data0_type
If item0(no%).data(0).value = "" Or item0(no%).data(0).sig <> "*" Or _
     item0(no%).data(0).poi(2) <= 0 Then
   Exit Function
End If
For i% = 0 To 1
 k% = (i% + 1) Mod 2
 t_l_v.poi(0) = item0(no%).data(0).poi(2 * i%)
 t_l_v.poi(1) = item0(no%).data(0).poi(2 * i% + 1)
 t_l_v.poi(2) = -1
 Call search_for_two_line_value(t_l_v, 0, tn(0), 1)
 t_l_v.poi(2) = 30000
 Call search_for_two_line_value(t_l_v, 0, tn(1), 1)
 last_n% = 0
 For j% = tn(0) + 1 To tn(1)
  n_% = two_line_value(j%).data(0).record.data1.index.i(0)
   If two_line_value(n_%).data(0).data0.para(0) = "1" And _
       (two_line_value(n_%).data(0).data0.para(1) = "1" Or _
          two_line_value(n_%).data(0).data0.para(1) = "-1") And _
            two_line_value(n_%).data(0).data0.poi(2) = item0(no%).data(0).poi(2 * k%) And _
             two_line_value(n_%).data(0).data0.poi(2) = item0(no%).data(0).poi(2 * k% + 1) Then
     last_n% = last_n% + 1
 ReDim Preserve n(last_n%) As Integer
 n(last_n%) = no%
 End If
 Next j%
 For j% = 1 To last_n%
 n_% = n(j%)
 combine_item0_with_two_line_value = combine_item0_with_two_line_value0(no%, n_%)
 If combine_item0_with_two_line_value > 1 Then
 Exit Function
 End If
 Next j%
Next i%
End Function
Public Function combine_item0_with_two_line_value0(ByVal no%, ByVal tlv%) As Byte
Dim ts(2) As String
Dim s(1) As String
Dim l(1) As Single
Dim temp_record As total_record_type
ts(0) = "1"
ts(1) = time_string("-1", two_line_value(tlv).data(0).data0.value, True, False)
ts(2) = item0(no%).data(0).value
Call solut_2order_equation(ts(0), ts(1), ts(2), s(0), s(1), False)
temp_record.record_data.data0.condition_data = item0(no%).data(0).record_for_value.data0.condition_data
Call add_conditions_to_record(two_line_value_, tlv%, 0, 0, temp_record.record_data.data0.condition_data)
temp_record.record_data.data0.theorem_no = -2
If two_line_value(tlv%).data(0).data0.para(1) = "-1" Then
 If Mid$(two_line_value(tlv%).data(0).data0.value, 1, 1) = "-" Then
  combine_item0_with_two_line_value0 = set_line_value(two_line_value(tlv%).data(0).data0.poi(2), _
     two_line_value(tlv%).data(0).data0.poi(3), s(0), two_line_value(tlv%).data(0).data0.n(2), _
      two_line_value(tlv%).data(0).data0.n(3), two_line_value(tlv%).data(0).data0.line_no(1), temp_record, 0, 0, False)
    If combine_item0_with_two_line_value0 > 1 Then
       Exit Function
    End If
 Else
  combine_item0_with_two_line_value0 = set_line_value(two_line_value(tlv%).data(0).data0.poi(0), _
     two_line_value(tlv%).data(0).data0.poi(1), s(0), two_line_value(tlv%).data(0).data0.n(0), _
       two_line_value(tlv%).data(0).data0.n(1), two_line_value(tlv%).data(0).data0.line_no(0), temp_record, 0, 0, False)
    If combine_item0_with_two_line_value0 > 1 Then
       Exit Function
    End If
 End If
ElseIf two_line_value(tlv%).data(0).data0.para(1) = "1" Then
 l(0) = squre_distance_point_point(m_poi(two_line_value(tlv%).data(0).data0.poi(0)).data(0).data0.coordinate, _
        m_poi(two_line_value(tlv%).data(0).data0.poi(1)).data(0).data0.coordinate)
 l(1) = squre_distance_point_point(m_poi(two_line_value(tlv%).data(0).data0.poi(2)).data(0).data0.coordinate, _
        m_poi(two_line_value(tlv%).data(0).data0.poi(3)).data(0).data0.coordinate)
 If l(0) > l(1) Then
  combine_item0_with_two_line_value0 = set_line_value(two_line_value(tlv%).data(0).data0.poi(2), _
     two_line_value(tlv%).data(0).data0.poi(3), s(1), two_line_value(tlv%).data(0).data0.n(2), _
       two_line_value(tlv%).data(0).data0.n(3), _
         two_line_value(tlv%).data(0).data0.line_no(1), temp_record, 0, 0, False)
    If combine_item0_with_two_line_value0 > 1 Then
       Exit Function
    End If
   combine_item0_with_two_line_value0 = set_line_value(two_line_value(tlv%).data(0).data0.poi(0), _
     two_line_value(tlv%).data(0).data0.poi(1), s(0), two_line_value(tlv%).data(0).data0.n(0), _
      two_line_value(tlv%).data(0).data0.n(1), two_line_value(tlv%).data(0).data0.line_no(0), temp_record, 0, 0, False)
    If combine_item0_with_two_line_value0 > 1 Then
       Exit Function
    End If
Else
  combine_item0_with_two_line_value0 = set_line_value(two_line_value(tlv%).data(0).data0.poi(2), _
     two_line_value(tlv%).data(0).data0.poi(3), s(0), two_line_value(tlv%).data(0).data0.n(2), _
      two_line_value(tlv%).data(0).data0.n(3), two_line_value(tlv%).data(0).data0.line_no(1), temp_record, 0, 0, False)
    If combine_item0_with_two_line_value0 > 1 Then
       Exit Function
    End If
   combine_item0_with_two_line_value0 = set_line_value(two_line_value(tlv%).data(0).data0.poi(0), _
     two_line_value(tlv%).data(0).data0.poi(1), s(0), two_line_value(tlv%).data(0).data0.n(0), _
      two_line_value(tlv%).data(0).data0.n(1), two_line_value(tlv%).data(0).data0.line_no(0), temp_record, 0, 0, False)
    If combine_item0_with_two_line_value0 > 1 Then
       Exit Function
    End If
 End If
End If
End Function
Public Function combine_item0_value_with_two_line_value(ByVal tlv%, ByVal it%) As Byte
If it% > 0 Then
   If item0(it%).data(0).sig = "*" Then
      If item0(it%).data(0).value <> "" Then
       For tlv% = 1 To last_conditions.last_cond(1).two_line_value_no
          If two_line_value(tlv%).data(0).data0.para(0) = "1" And _
              (two_line_value(tlv%).data(0).data0.para(1) = "1" Or _
                  two_line_value(tlv%).data(0).data0.para(0) = "-1") Then
             If two_line_value(tlv%).data(0).data0.poi(0) = item0(it%).data(0).poi(0) And _
                 two_line_value(tlv%).data(0).data0.poi(1) = item0(it%).data(0).poi(1) And _
                  two_line_value(tlv%).data(0).data0.poi(2) = item0(it%).data(0).poi(2) And _
                   two_line_value(tlv%).data(0).data0.poi(3) = item0(it%).data(0).poi(3) Then
              combine_item0_value_with_two_line_value = _
                 combine_item0_value_with_two_line_value0(ByVal it%, ByVal tlv%)
                  If combine_item0_value_with_two_line_value > 1 Then
                      Exit Function
                  End If
             End If
          End If
       Next tlv%
      End If
   End If
Else
    If two_line_value(tlv%).data(0).data0.para(0) = "1" And _
              (two_line_value(tlv%).data(0).data0.para(1) = "1" Or _
                  two_line_value(tlv%).data(0).data0.para(0) = "-1") Then
       For it% = 1 To last_conditions.last_cond(1).item0_no
          If item0(it%).data(0).sig = "*" Then
             If two_line_value(tlv%).data(0).data0.poi(0) = item0(it%).data(0).poi(0) And _
                 two_line_value(tlv%).data(0).data0.poi(1) = item0(it%).data(0).poi(1) And _
                  two_line_value(tlv%).data(0).data0.poi(2) = item0(it%).data(0).poi(2) And _
                   two_line_value(tlv%).data(0).data0.poi(3) = item0(it%).data(0).poi(3) Then
              combine_item0_value_with_two_line_value = _
                 combine_item0_value_with_two_line_value0(ByVal it%, ByVal tlv%)
                  If combine_item0_value_with_two_line_value > 1 Then
                      Exit Function
                  End If
             End If
             
          End If
       Next it%
    End If

End If
End Function
Public Function combine_item0_value_with_two_line_value0(ByVal it%, ByVal tlv%) As Byte
Dim temp_record As total_record_type
Dim tv(1) As String
Call add_conditions_to_record(two_line_value_, tlv%, 0, 0, temp_record.record_data.data0.condition_data)
Call add_record_to_record(item0(it%).data(0).record_for_value.data0.condition_data, _
                                              temp_record.record_data.data0.condition_data)
temp_record.record_data.data0.theorem_no = 1
If two_line_value(tlv%).data(0).data0.para(1) = "1" Then
   If solut_2order_equation("1", time_string(two_line_value(tlv%).data(0).data0.value, "-1", _
          True, False), item0(it%).data(0).value, tv(0), tv(1), False) Then
           If tv(0) <> "F" And tv(1) <> "F" Then
                If (m_poi(item0(it%).data(0).poi(0)).data(0).data0.coordinate.X - _
                           m_poi(item0(it%).data(0).poi(1)).data(0).data0.coordinate.X) ^ 2 + _
                              (m_poi(item0(it%).data(0).poi(0)).data(0).data0.coordinate.Y - _
                                m_poi(item0(it%).data(0).poi(1)).data(0).data0.coordinate.Y) ^ 2 > _
                              (m_poi(item0(it%).data(0).poi(2)).data(0).data0.coordinate.X - _
                                m_poi(item0(it%).data(0).poi(3)).data(0).data0.coordinate.X) ^ 2 + _
                              (m_poi(item0(it%).data(0).poi(2)).data(0).data0.coordinate.Y - _
                                m_poi(item0(it%).data(0).poi(3)).data(0).data0.coordinate.Y) ^ 2 Then
                 combine_item0_value_with_two_line_value0 = set_line_value( _
                   item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), tv(0), _
                    item0(it%).data(0).n(0), item0(it%).data(0).n(1), item0(it%).data(0).line_no(0), _
                      temp_record, 0, 0, False)
                      If combine_item0_value_with_two_line_value0 > 1 Then
                         Exit Function
                      End If
                 combine_item0_value_with_two_line_value0 = set_line_value( _
                   item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), tv(1), _
                    item0(it%).data(0).n(2), item0(it%).data(0).n(3), item0(it%).data(0).line_no(1), _
                      temp_record, 0, 0, False)
                      If combine_item0_value_with_two_line_value0 > 1 Then
                         Exit Function
                      End If
                Else
                 combine_item0_value_with_two_line_value0 = set_line_value( _
                   item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), tv(1), _
                    item0(it%).data(0).n(0), item0(it%).data(0).n(1), item0(it%).data(0).line_no(0), _
                      temp_record, 0, 0, False)
                      If combine_item0_value_with_two_line_value0 > 1 Then
                         Exit Function
                      End If
                 combine_item0_value_with_two_line_value0 = set_line_value( _
                   item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), tv(0), _
                    item0(it%).data(0).n(2), item0(it%).data(0).n(3), item0(it%).data(0).line_no(1), _
                      temp_record, 0, 0, False)
                      If combine_item0_value_with_two_line_value0 > 1 Then
                         Exit Function
                      End If
                End If
           End If
      End If
Else
   If solut_2order_equation("1", time_string(two_line_value(tlv%).data(0).data0.value, "-1", _
          True, False), time_string("-1", item0(it%).data(0).value, True, False), tv(1), tv(0), False) Then
           tv(0) = time_string(tv(0), "-1", True, False)
           If tv(0) <> "F" And tv(1) <> "F" Then
                If (m_poi(item0(it%).data(0).poi(0)).data(0).data0.coordinate.X - _
                           m_poi(item0(it%).data(0).poi(1)).data(0).data0.coordinate.X) ^ 2 + _
                              (m_poi(item0(it%).data(0).poi(0)).data(0).data0.coordinate.Y - _
                                m_poi(item0(it%).data(0).poi(1)).data(0).data0.coordinate.Y) ^ 2 > _
                              (m_poi(item0(it%).data(0).poi(2)).data(0).data0.coordinate.X - _
                                m_poi(item0(it%).data(0).poi(3)).data(0).data0.coordinate.X) ^ 2 + _
                              (m_poi(item0(it%).data(0).poi(2)).data(0).data0.coordinate.Y - _
                                m_poi(item0(it%).data(0).poi(3)).data(0).data0.coordinate.Y) ^ 2 Then
                 combine_item0_value_with_two_line_value0 = set_line_value( _
                   item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), tv(0), _
                    item0(it%).data(0).n(0), item0(it%).data(0).n(1), item0(it%).data(0).line_no(0), _
                      temp_record, 0, 0, False)
                      If combine_item0_value_with_two_line_value0 > 1 Then
                         Exit Function
                      End If
                 combine_item0_value_with_two_line_value0 = set_line_value( _
                   item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), tv(1), _
                    item0(it%).data(0).n(2), item0(it%).data(0).n(3), item0(it%).data(0).line_no(1), _
                      temp_record, 0, 0, False)
                      If combine_item0_value_with_two_line_value0 > 1 Then
                         Exit Function
                      End If
                Else
                 combine_item0_value_with_two_line_value0 = set_line_value( _
                   item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), tv(1), _
                    item0(it%).data(0).n(0), item0(it%).data(0).n(1), item0(it%).data(0).line_no(0), _
                      temp_record, 0, 0, False)
                      If combine_item0_value_with_two_line_value0 > 1 Then
                         Exit Function
                      End If
                 combine_item0_value_with_two_line_value0 = set_line_value( _
                   item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), tv(0), _
                    item0(it%).data(0).n(2), item0(it%).data(0).n(3), item0(it%).data(0).line_no(1), _
                      temp_record, 0, 0, False)
                      If combine_item0_value_with_two_line_value0 > 1 Then
                         Exit Function
                      End If
                End If
           End If
      End If
End If
End Function
Public Function combine_two_line_value_with_item(ByVal tlv%) As Byte
Dim i%, j%, k%, no%, j_%
Dim last_tn%
Dim n_(1) As Integer
Dim tn() As Integer
Dim it As item0_data_type
For i% = 0 To 1
 For j% = 0 To 1
      j_% = (j% + 1) Mod 2
  it.poi(2 * j%) = two_line_value(tlv%).data(0).data0.poi(2 * i%)
   it.poi(2 * j% + 1) = two_line_value(tlv%).data(0).data0.poi(2 * i% + 1)
  it.poi(2 * j_%) = -1
  Call search_for_item0(it, j%, n_(0), 1)
  it.poi(2 * j_%) = 30000
  Call search_for_item0(it, j%, n_(1), 1)
last_tn% = 0
For k% = n_(0) + 1 To n_(1)
 no% = item0(k%).data(0).index(j%)
    last_tn% = last_tn% + 1
    ReDim Preserve tn(last_tn%) As Integer
    tn(last_tn%) = no%
Next k%
For k% = 1 To last_tn%
 no% = tn(k%)
  combine_two_line_value_with_item = combine_item0_with_two_line_value_(no%, tlv%, j%, i%)
  If combine_two_line_value_with_item > 1 Then
   Exit Function
  End If
 Next k%
Next j%
Next i%
End Function

Public Function combine_relation_with_h_point_pair0(v$, dp As point_pair_data0_type, ByVal k1%, ByVal k2%, ByVal k3%, ByVal k4%, _
                                     re As total_record_type) As Byte
'调和比
Dim ts$
Dim temp_record As total_record_type
If dp.is_h_ratio = 1 Then
 If is_same_two_point(k1%, k2%, 0, 2) Or is_same_two_point(k1%, k2%, 1, 3) Then
  If (k1% = 2 And k2% = 0) Or (k1% = 1 And k1% = 3) Then
   Call exchange_two_integer(k1%, k2%)
   Call exchange_two_integer(k3%, k4%)
   v$ = divide_string("1", v$, True, False)
  End If
  temp_record = re
  ts$ = divide_string(minus_string("1", v$, False, False), add_string("1", v$, False, False), True, False)
  combine_relation_with_h_point_pair0 = set_Drelation(dp.poi(2 * k3%), dp.poi(2 * k3% + 1), dp.poi(2 * k4%), _
        dp.poi(2 * k4% + 1), dp.n(2 * k3%), dp.n(2 * k3% + 1), dp.n(2 * k4%), dp.n(2 * k4% + 1), dp.line_no(k3%), dp.line_no(k4%), _
         ts$, temp_record, 0, 0, 0, 0, 0, False)
 End If
End If
End Function
Public Function combine_area_relation_with_area_relation_(tri_re1%, tri_re2%, j%, k%) As Byte
Dim temp_record As total_record_type
Dim triA1_(2) As condition_type
Dim triA2_(2) As condition_type
Dim v1(1) As String
Dim v2(1) As String
'On Error GoTo combine_area_relation_ith_area_relation_error
 temp_record.record_data.data0.condition_data.condition(1).ty = area_relation_
  temp_record.record_data.data0.condition_data.condition(2).ty = area_relation_
   temp_record.record_data.data0.condition_data.condition(1).no = tri_re1%
    temp_record.record_data.data0.condition_data.condition(2).no = tri_re2%
      temp_record.record_data.data0.condition_data.condition_no = 2
If Darea_relation(tri_re1%).data(0).area_element(2).no > 0 Then
  Call read_ratio_from_relation(Darea_relation(tri_re1%).data(0).value, _
         j%, v1(0), v1(1), True, 3)
 triA1_(0) = Darea_relation(tri_re1%).data(0).area_element((j% + 1) Mod 3)
 triA1_(1) = Darea_relation(tri_re1%).data(0).area_element((j% + 2) Mod 3)
Else
   Call read_ratio_from_relation(Darea_relation(tri_re1%).data(0).value, _
         j%, v1(0), v1(1), True, 0)
 triA1_(0) = Darea_relation(tri_re1%).data(0).area_element((j% + 1) Mod 2)
 triA1_(1).no = 0
 End If
If Darea_relation(tri_re2%).data(0).area_element(2).no > 0 Then
  Call read_ratio_from_relation(Darea_relation(tri_re2%).data(0).value, _
         k%, v2(0), v2(1), True, 3)
 triA2_(0) = Darea_relation(tri_re2%).data(0).area_element((k% + 1) Mod 3)
 triA2_(1) = Darea_relation(tri_re2%).data(0).area_element((k% + 2) Mod 3)
 Else
   Call read_ratio_from_relation(Darea_relation(tri_re2%).data(0).value, _
         k%, v2(0), v2(1), True, 0)
 triA2_(0) = Darea_relation(tri_re2%).data(0).area_element((k% + 1) Mod 2)
 triA2_(1).no = 0
 End If
If v1(0) <> "" And v2(0) <> "" Then
 combine_area_relation_with_area_relation_ = _
      set_area_relation(triA1_(0), triA2_(0), divide_string(v2(0), v1(0), True, False), _
        temp_record, 0, 0, 0)
If combine_area_relation_with_area_relation_ > 1 Then
         Exit Function
End If
End If
If v1(1) <> "" And v2(0) <> "" Then
 combine_area_relation_with_area_relation_ = set_area_relation(triA1_(1), triA2_(0), divide_string(v2(0), v1(1), True, False), _
        temp_record, 0, 0, 0)
If combine_area_relation_with_area_relation_ > 1 Then
         Exit Function
End If
End If
If v1(0) <> "" And v2(1) <> "" Then
 combine_area_relation_with_area_relation_ = set_area_relation(triA1_(0), triA2_(1), divide_string(v2(1), v1(0), True, False), _
        temp_record, 0, 0, 0)
If combine_area_relation_with_area_relation_ > 1 Then
         Exit Function
End If
End If
If v1(1) <> "" And v2(1) <> "" Then
 combine_area_relation_with_area_relation_ = set_area_relation(triA1_(1), triA2_(1), divide_string(v2(1), v1(1), True, False), _
        temp_record, 0, 0, 0)
If combine_area_relation_with_area_relation_ > 1 Then
         Exit Function
End If
End If
combine_area_relation_ith_area_relation_error:
End Function
Public Function combine_area_relation_with_area_relation(tri_re%) As Byte
Dim i%, j%, k%, tn%
Dim tn_(1) As Integer
Dim triA_re As area_relation_data_type
Dim n() As Integer
Dim last_n As Integer
For i% = 0 To 2
 If Darea_relation(tri_re%).data(0).area_element(i%).no > 0 Then
  For j% = 0 To 2
   triA_re.area_element(j%) = Darea_relation(tri_re%).data(0).area_element(i%)
    triA_re.area_element((j% + 1) Mod 3).no = -1
     triA_re.area_element((j% + 1) Mod 3).ty = 0
     Call search_for_area_relation(triA_re, j%, tn_(0), 1)
    triA_re.area_element((j% + 1) Mod 3).no = 30000
     triA_re.area_element((j% + 1) Mod 3).ty = 255
     Call search_for_area_relation(triA_re, j%, tn_(1), 1)
  last_n% = 0
  For k% = tn_(0) + 1 To tn_(1)
   tn% = Darea_relation(k%).data(0).record.data1.index.i(j%)
   If tn% < tri_re% Then
    last_n% = last_n% + 1
     ReDim Preserve n(last_n%) As Integer
      n(last_n%) = tn%
   End If
  Next k%
  For k% = 1 To last_n%
   combine_area_relation_with_area_relation = _
    combine_area_relation_with_area_relation_(tri_re%, n(k%), i%, j%)
     If combine_area_relation_with_area_relation > 1 Then
         Exit Function
     End If
  Next k%
 Next j%
 End If
Next i%
End Function
Public Function combine_two_area_elemenet(area_el1 As condition_type, _
                 area_el2 As condition_type, out_area_el1 As condition_type, _
                  out_area_el2 As condition_type, out_area_el3 As condition_type) As Byte
Dim tn%
If area_el1.ty = triangle_ And area_el2.ty = triangle_ Then
   combine_two_area_elemenet = combine_two_triangle(area_el1.no, area_el2.no, _
       out_area_el1.no, out_area_el2.no, out_area_el3.no, tn%)
      out_area_el1.ty = triangle_
      out_area_el2.ty = triangle_
      out_area_el3.ty = triangle_
   If tn% > 0 Then
     out_area_el3.ty = polygon_
     out_area_el3.no = tn%
   End If
ElseIf area_el1.ty = polygon_ And area_el2.ty = triangle_ Then
ElseIf area_el1.ty = triangle_ And area_el2.ty = polygon_ Then
ElseIf area_el1.ty = polygon_ And area_el2.ty = polygon_ Then
End If
End Function
Public Function combine_triangle_with_polygon(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p5%, _
                    ByVal p6%, ByVal p7%, ByVal p8%, ByVal n1%, ByVal n2%, ByVal n3%, ByVal n4%, _
                      po1 As polygon, po2 As polygon) As Integer
'p2%=p6%,p3%=p7% p1%(n1%),p2%(n2%),p5%(n3%),p6%(n%)
'=1 p1+p2=po1
'=2 p1
    'p2%=p6%,p3%=p7%'1 v1+v2=po.-1 v1-v2=po ,-2 v2-v1=po,-3 v1-v2%=po1+po2,-4,v2-v1=po1+po2,
    '2,v1+v2%=po2-po1,3  ,v1+v2%=po2+po1 -5 v1-v2=op1-po2,-6 v1-v2=op2-po1
Dim tA(3) As Integer
Dim tp%
    If (n1% < n2% And n3% < n4% And n2% < n4%) Or _
            (n1% > n2% And n3% > n4% And n2% > n4%) Then 'v2>v1
            po1.total_v = 4
              po1.v(0) = p3%
               po1.v(1) = p1%
                po1.v(2) = p5%
                 po1.v(3) = p8%
                  combine_triangle_with_polygon = -2
                   Exit Function    ' 三角形含在四边形中
    ElseIf (n1% < n2% And n3% < n4% And n2% > n4%) Or _
               (n1% > n2% And n3% > n4% And n2% < n4%) Then ' 三角形顶点在四边形外(右)
      tA(0) = angle_number(p2%, p1%, p3%, "", 0)
      tA(1) = angle_number(p3%, p1%, p8%, "", 0)
      If (tA(0) > 0 And tA(1) > 0) Or (tA(0) < 0 And tA(1) < 0) Then ' 四边形顶点在三角形外
            po1.total_v = 4
              po1.v(0) = p1%
               po1.v(1) = p5%
                po1.v(2) = p3%
                 po1.v(2) = p8%
                 combine_triangle_with_polygon = -5
                   Exit Function
      ElseIf (tA(0) > 0 And tA(1) < 0) Or (tA(0) < 0 And tA(1) > 0) Then ' 四边形顶点在三角形内
             po1.total_v = 3
              po1.v(0) = p1%
               po1.v(1) = p5%
                po1.v(2) = p8%
            po2.total_v = 3
              po2.v(0) = p3%
               po2.v(1) = p1%
                po2.v(2) = p8%
                 combine_triangle_with_polygon = -3
                  Exit Function
         End If
    ElseIf (n1% < n2% And n3% > n4%) Or _
               (n1% > n2% And n3% < n4%) Then    ' 三角形顶点在四边形外(左)
        tA(0) = angle_number(p3%, p1%, p2%, "", 0)
        tA(1) = angle_number(p8%, p1%, p3%, "", 0)
        If (tA(0) > 0 And tA(1) > 0) Or (tA(0) < 0 And tA(1) < 0) Then
            po1.total_v = 3
              po1.v(0) = p1%
               po1.v(1) = p5%
                po1.v(2) = p8%
            po2.total_v = 3
              po2.v(0) = p3%
               po2.v(1) = p1%
                po2.v(2) = p8%
                 combine_triangle_with_polygon = 2
                  Exit Function
        ElseIf (tA(0) > 0 And tA(1) < 0) Or (tA(0) < 0 And tA(1) > 0) Then
            po1.total_v = 4
              po1.v(0) = p1%
               po1.v(1) = p5%
                po1.v(2) = p8%
                 po1.v(3) = p3%
                  combine_triangle_with_polygon = 1
                   Exit Function
        End If
      End If
End Function
Public Function combine_two_polygon0_with_3point(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
                  ByVal p5%, ByVal p6%, ByVal p7%, ByVal p8%, po1 As polygon, _
                    po2 As polygon) As Integer
'p1%=p5%,p2%=p6%,p3%=p7%,p2(n1%),p1(n2),p3(n3),p4(n4),p6(n5),p5(n6),p7(n7),p8(n8)
Dim tA(3) As Integer
Dim tp%
     tA(0) = angle_number(p4%, p3%, p2%, "", 0)
     tA(1) = angle_number(p8%, p3%, p4%, "", 0)
     tA(2) = angle_number(p4%, p1%, p2%, "", 0)
     tA(3) = angle_number(p8%, p1%, p4%, "", 0)
        If (tA(0) > 0 And tA(1) > 0) Or (tA(0) < 0 And tA(1)) < 0 Then
            tA(0) = 1
        Else
            tA(0) = -1
        End If
        If (tA(2) > 0 And tA(3) > 0) Or (tA(2) < 0 And tA(3)) < 0 Then
            tA(2) = 1
        Else
            tA(2) = -1
        End If
     If tA(0) > 0 And tA(2) > 0 Then 'V1<v2
        po1.total_v = 3
        po1.v(0) = p3%
        po1.v(1) = p4%
        po1.v(2) = p8%
        po2.total_v = 3
        po2.v(0) = p1%
        po2.v(1) = p4%
        po2.v(2) = p8%
        combine_two_polygon0_with_3point = -4
         Exit Function
     ElseIf (tA(0) < 0 And tA(2) < 0) Or (tA(0) < 0 And tA(2) > 0) Then
        po1.total_v = 3
        po1.v(0) = p4%
        po1.v(1) = p1%
        po1.v(2) = p3%
        po1.v(3) = p8%
        combine_two_polygon0_with_3point = -5
         Exit Function
     End If
End Function
Public Function combine_two_polygon0_with_3line(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
          ByVal p5%, ByVal p6%, ByVal p7%, ByVal p8%, ByVal n1%, ByVal n2%, ByVal n3%, _
            ByVal n4%, ByVal n5%, ByVal n6%, ByVal n7%, ByVal n8%, po1 As polygon, _
              po2 As polygon) As Integer
'p1%=p5%,p2%=p6%,p3%=p7%,p2(n1%),p1(n2),p3(n3),p4(n4),p6(n5),p5(n6),p7(n7),p8(n8)
Dim ty(1) As Byte '1 和 2 前大 3后大
Dim tA(3) As Integer
Dim tp%
If (n1% < n2% And n5% < n6% And n2% < n6%) Or (n1% > n2% And n5% > n6% And n2% > n6%) Then
 ty(0) = 3
ElseIf (n1% < n2% And n5% < n6% And n2% > n6%) Or (n1% > n2% And n5% > n6% And n2% < n6%) Then
 ty(0) = 2
ElseIf (n1% < n2% And n5% > n6%) Or (n1% > n2% And n5% < n6%) Then
 ty(0) = 1
End If
If (n3% < n4% And n7% < n8% And n4% < n8%) Or (n3% > n4% And n7% > n8% And n4% > n8%) Then
 ty(1) = 3
ElseIf (n3% < n4% And n7% < n8% And n4% > n8%) Or (n3% > n4% And n7% > n8% And n4% < n8%) Then
 ty(1) = 2
ElseIf (n3% < n4% And n7% > n8%) Or (n3% > n4% And n7% < n8%) Then
 ty(1) = 1
End If
If ty(0) = 1 And ty(1) = 1 Then '
 po1.total_v = 4
 po1.v(0) = p1%
 po1.v(1) = p5%
 po1.v(2) = p8%
 po1.v(4) = p4%
 combine_two_polygon0_with_3line = 1
 Exit Function
ElseIf ty(0) = 2 And ty(1) = 2 Then 'v1>v2
  po1.total_v = 4
 po1.v(0) = p1%
 po1.v(1) = p5%
 po1.v(2) = p8%
 po1.v(4) = p4%
 combine_two_polygon0_with_3line = -1
ElseIf ty(0) = 2 And ty(1) = 3 Then
 po1.total_v = 4
 po1.v(0) = p4%
 po1.v(1) = p3%
 po1.v(2) = p8%
 po1.v(3) = p7%
 combine_two_polygon0_with_3line = -5
 Exit Function
ElseIf ty(0) = 3 And ty(1) = 3 Then
  po1.total_v = 4
 po1.v(0) = p1%
 po1.v(1) = p5%
 po1.v(2) = p8%
 po1.v(4) = p4%
 combine_two_polygon0_with_3line = -2
ElseIf ty(0) = 3 And ty(1) = 2 Then
 po1.total_v = 4
 po1.v(0) = p4%
 po1.v(1) = p3%
 po1.v(2) = p8%
 po1.v(3) = p7%
 combine_two_polygon0_with_3line = -5
 Exit Function
End If
End Function
Public Function combine_two_area_of_element0(A_ele1 As condition_type, _
          A_ele2 As condition_type, o_ele1 As condition_type, o_ele2 As condition_type, _
            relation As relation_data0_type) As Integer
Dim pol(3) As polygon
Dim no%, tp%
Dim A_element(1) As condition_type
Dim v$
If A_ele1.ty = triangle_ And A_ele1.no > 0 Then
 pol(0).total_v = 3
 pol(0).v(0) = triangle(A_ele1.no).data(0).poi(0)
 pol(0).v(1) = triangle(A_ele1.no).data(0).poi(1)
 pol(0).v(2) = triangle(A_ele1.no).data(0).poi(2)
Else
 pol(0).total_v = 4
 pol(0).v(0) = Dpolygon4(A_ele1.no).data(0).poi(0)
 pol(0).v(1) = Dpolygon4(A_ele1.no).data(0).poi(1)
 pol(0).v(2) = Dpolygon4(A_ele1.no).data(0).poi(2)
 pol(0).v(3) = Dpolygon4(A_ele1.no).data(0).poi(3)
End If
If A_ele2.ty = triangle_ Then
 pol(1).total_v = 3
 pol(1).v(0) = triangle(A_ele2.no).data(0).poi(0)
 pol(1).v(1) = triangle(A_ele2.no).data(0).poi(1)
 pol(1).v(2) = triangle(A_ele2.no).data(0).poi(2)
Else
 pol(1).total_v = 4
 pol(1).v(0) = Dpolygon4(A_ele2.no).data(0).poi(0)
 pol(1).v(1) = Dpolygon4(A_ele2.no).data(0).poi(1)
 pol(1).v(2) = Dpolygon4(A_ele2.no).data(0).poi(2)
 pol(1).v(3) = Dpolygon4(A_ele2.no).data(0).poi(3)
End If
combine_two_area_of_element0 = combine_two_polygon(pol(0), pol(1), pol(2), pol(3), relation)
 o_ele1 = change_poly_to_area_element(pol(2))
 o_ele2 = change_poly_to_area_element(pol(3))
End Function

Public Function combine_two_area_of_element(A_ele1 As area_of_element_data_type, _
          A_ele2 As area_of_element_data_type, re As total_record_type) As Byte
Dim temp_record As total_record_type
Dim pol(3) As polygon
Dim triA(3) As Integer
Dim ty As Integer
Dim no%, tp%
Dim A_element(1) As condition_type
Dim v$
Dim relation As relation_data0_type
temp_record = re
temp_record.record_data.data0.theorem_no = 1
'If A_ele1.element.ty = triangle_ Then
' pol(0).total_v = 3
' pol(0).v(0) = triangle(A_ele1.element.no).data(0).poi(0)
' pol(0).v(1) = triangle(A_ele1.element.no).data(0).poi(1)
' pol(0).v(2) = triangle(A_ele1.element.no).data(0).poi(2)
'Else
' pol(0).total_v = 4
' pol(0).v(0) = Dpolygon4(A_ele1.element.no).data(0).poi(0)
' pol(0).v(1) = Dpolygon4(A_ele1.element.no).data(0).poi(1)
' pol(0).v(2) = Dpolygon4(A_ele1.element.no).data(0).poi(2)
' pol(0).v(3) = Dpolygon4(A_ele1.element.no).data(0).poi(3)
'End If
'If A_ele2.element.ty = triangle_ Then
' pol(1).total_v = 3
' pol(1).v(0) = triangle(A_ele2.element.no).data(0).poi(0)
' pol(1).v(1) = triangle(A_ele2.element.no).data(0).poi(1)
' pol(1).v(2) = triangle(A_ele2.element.no).data(0).poi(2)
'Else
' pol(1).total_v = 4
' pol(1).v(0) = Dpolygon4(A_ele2.element.no).data(0).poi(0)
' pol(1).v(1) = Dpolygon4(A_ele2.element.no).data(0).poi(1)
' pol(1).v(2) = Dpolygon4(A_ele2.element.no).data(0).poi(2)
' pol(1).v(3) = Dpolygon4(A_ele2.element.no).data(0).poi(3)
'End If
ty = combine_two_area_of_element0(A_ele1.element, A_ele2.element, A_element(0), A_element(1), relation)
If A_element(0).ty = triangle_ Then
   pol(2).total_v = 3
   pol(2).v(0) = triangle(A_element(0).no).data(0).poi(0)
   pol(2).v(1) = triangle(A_element(0).no).data(0).poi(1)
   pol(2).v(2) = triangle(A_element(0).no).data(0).poi(2)
Else
   pol(2).total_v = 4
   pol(2).v(0) = Dpolygon4(A_element(0).no).data(0).poi(0)
   pol(2).v(1) = Dpolygon4(A_element(0).no).data(0).poi(1)
   pol(2).v(2) = Dpolygon4(A_element(0).no).data(0).poi(2)
   pol(2).v(3) = Dpolygon4(A_element(0).no).data(0).poi(3)
End If
If A_element(1).ty = triangle_ Then
   pol(3).total_v = 3
   pol(3).v(0) = triangle(A_element(1).no).data(0).poi(0)
   pol(3).v(1) = triangle(A_element(1).no).data(0).poi(1)
   pol(3).v(2) = triangle(A_element(1).no).data(0).poi(2)
Else
   pol(3).total_v = 4
   pol(3).v(0) = Dpolygon4(A_element(1).no).data(0).poi(0)
   pol(3).v(1) = Dpolygon4(A_element(1).no).data(0).poi(1)
   pol(3).v(2) = Dpolygon4(A_element(1).no).data(0).poi(2)
   pol(3).v(3) = Dpolygon4(A_element(1).no).data(0).poi(3)
End If
'ty = combine_two_polygon(pol(0), pol(1), pol(2), pol(3), relation)
'temp_record = re
If relation.ty > 0 And th_chose(157).chose = 1 Then
   temp_record.record_data.data0.theorem_no = 157
    v$ = divide_string(A_ele1.value, A_ele2.value, True, False)
       Call ratio_value1(v$, relation.ty, relation.value)
   combine_two_area_of_element = set_Drelation(relation.poi(0), relation.poi(1), _
    relation.poi(2), relation.poi(3), relation.n(0), relation.n(1), relation.n(2), _
     relation.n(3), relation.line_no(0), relation.line_no(1), relation.value, temp_record, _
      0, 0, 0, 0, 0, False)
      If combine_two_area_of_element > 1 Then
         Exit Function
      End If
End If
If ty = 3 Then
    v$ = add_string(A_ele1.value, A_ele2.value, True, False)
' A_element(0) = change_poly_to_area_element(pol(2))
' A_element(1) = change_poly_to_area_element(pol(3))
   If is_area_of_element(A_element(0).ty, A_element(0).no, no%, -1000) Then
      v$ = minus_string(v$, area_of_element(no%).data(0).value, True, False)
       Call add_conditions_to_record(area_of_element_, no%, 0, 0, temp_record.record_data.data0.condition_data)
       combine_two_area_of_element = set_area_of_element(A_element(1).ty, A_element(1).no, _
             v$, 0, temp_record)
       If combine_two_area_of_element > 1 Then
          Exit Function
       End If
   ElseIf is_area_of_element(A_element(1).ty, A_element(1).no, no%, -1000) Then
      v$ = minus_string(v$, area_of_element(no%).data(1).value, True, False)
       combine_two_area_of_element = set_area_of_element(A_element(0).ty, A_element(0).no, _
             v$, 0, temp_record)
       If combine_two_area_of_element > 1 Then
          Exit Function
       End If
   End If
ElseIf ty = 2 Then
    v$ = add_string(A_ele1.value, A_ele2.value, True, False)
 'A_element(0) = change_poly_to_area_element(pol(2))
 'A_element(1) = change_poly_to_area_element(pol(3))
   If is_area_of_element(A_element(0).ty, A_element(0).no, no%, -1000) Then
      v$ = minus_string(area_of_element(no%).data(0).value, v$, True, False)
       Call add_conditions_to_record(area_of_element_, no%, 0, 0, temp_record.record_data.data0.condition_data)
       combine_two_area_of_element = set_area_of_element(A_element(1).ty, A_element(1).no, _
             v$, 0, temp_record)
       If combine_two_area_of_element > 1 Then
          Exit Function
       End If
   ElseIf is_area_of_element(A_element(1).ty, A_element(1).no, no%, -1000) Then
      v$ = add_string(area_of_element(no%).data(0).value, v$, True, False)
       combine_two_area_of_element = set_area_of_element(A_element(0).ty, A_element(0).no, _
             v$, 0, temp_record)
       If combine_two_area_of_element > 1 Then
          Exit Function
       End If
   End If
ElseIf ty = 1 Then
temp_record = re
temp_record.record_data.data0.theorem_no = 1
 If pol(2).total_v = 3 Then
 combine_two_area_of_element = set_area_of_element(triangle_, _
      triangle_number(pol(2).v(0), pol(2).v(1), pol(2).v(2), 0, 0, 0, 0, 0, 0, 0), _
       add_string(A_ele1.value, A_ele2.value, True, False), 0, temp_record)
 Else
 combine_two_area_of_element = set_area_of_element(polygon_, _
      polygon4_number(pol(2).v(0), pol(2).v(1), pol(2).v(2), pol(2).v(3), 0), _
       add_string(A_ele1.value, A_ele2.value, True, False), 0, temp_record)
 End If
ElseIf ty = -1 Then
 If pol(2).total_v = 3 Then
 combine_two_area_of_element = set_area_of_element(triangle_, _
      triangle_number(pol(2).v(0), pol(2).v(1), pol(2).v(2), 0, 0, 0, 0, 0, 0, 0), _
       minus_string(A_ele1.value, A_ele2.value, True, False), 0, temp_record)
 Else
 combine_two_area_of_element = set_area_of_element(polygon_, _
      polygon4_number(pol(2).v(0), pol(2).v(1), pol(2).v(2), pol(2).v(3), 0), _
       minus_string(A_ele1.value, A_ele2.value, True, False), 0, temp_record)
 End If
ElseIf ty = -2 Then
 If pol(2).total_v = 3 Then
 combine_two_area_of_element = set_area_of_element(triangle_, _
      triangle_number(pol(2).v(0), pol(2).v(1), pol(2).v(2), 0, 0, 0, 0, 0, 0, 0), _
       minus_string(A_ele2.value, A_ele1.value, True, False), 0, temp_record)
 Else
 combine_two_area_of_element = set_area_of_element(polygon_, _
      polygon4_number(pol(2).v(0), pol(2).v(1), pol(2).v(2), pol(2).v(3), 0), _
       minus_string(A_ele2.value, A_ele1.value, True, False), 0, temp_record)
 End If
ElseIf ty = -3 Then
 'A_element(0) = change_poly_to_area_element(pol(2))
 'A_element(1) = change_poly_to_area_element(pol(3))
 v$ = minus_string(A_ele1.value, A_ele2.value, False, False)
 If is_area_of_element0(A_element(0), no%, -1000) Then
   Call add_conditions_to_record(area_of_element_, no%, 0, 0, temp_record.record_data.data0.condition_data)
  v$ = minus_string(v$, area_of_element(no%).data(0).value, True, False)
       Call add_conditions_to_record(area_of_element_, no%, 0, 0, temp_record.record_data.data0.condition_data)
   combine_two_area_of_element = set_area_of_element(A_element(1).ty, _
      A_element(1).no, v$, 0, temp_record)
    If combine_two_area_of_element > 1 Then
     Exit Function
    End If
  ElseIf is_area_of_element0(A_element(1), no%, -1000) Then
   Call add_conditions_to_record(area_of_element_, no%, 0, 0, temp_record.record_data.data0.condition_data)
    v$ = minus_string(v$, area_of_element(no%).data(0).value, True, False)
       Call add_conditions_to_record(area_of_element_, no%, 0, 0, temp_record.record_data.data0.condition_data)
    combine_two_area_of_element = set_area_of_element(A_element(0).ty, _
      A_element(1).no, v$, 0, temp_record)
   If combine_two_area_of_element > 1 Then
    Exit Function
   End If
  End If
ElseIf ty = -4 Then
' A_element(0) = change_poly_to_area_element(pol(2))
'  A_element(1) = change_poly_to_area_element(pol(3))
   v$ = minus_string(A_ele2.value, A_ele1.value, False, False)
    If is_area_of_element0(A_element(0), no%, -1000) Then
     Call add_conditions_to_record(area_of_element_, no%, 0, 0, temp_record.record_data.data0.condition_data)
      v$ = minus_string(v$, area_of_element(no%).data(0).value, True, False)
     combine_two_area_of_element = set_area_of_element(A_element(1).ty, _
      A_element(1).no, v$, 0, temp_record)
    If combine_two_area_of_element > 1 Then
     Exit Function
    End If
  ElseIf is_area_of_element0(A_element(1), no%, -1000) Then
   Call add_conditions_to_record(area_of_element_, no%, 0, 0, temp_record.record_data.data0.condition_data)
    v$ = minus_string(v$, area_of_element(no%).data(0).value, True, False)
    combine_two_area_of_element = set_area_of_element(A_element(0).ty, _
      A_element(0).no, v$, 0, temp_record)
   If combine_two_area_of_element > 1 Then
    Exit Function
   End If
  End If
ElseIf ty = -5 Then
triA(0) = triangle_number(pol(2).v(0), pol(2).v(1), pol(2).v(2), 0, 0, 0, 0, 0, 0, 0)
triA(1) = triangle_number(pol(2).v(3), pol(2).v(1), pol(2).v(2), 0, 0, 0, 0, 0, 0, 0)
triA(2) = triangle_number(pol(2).v(0), pol(2).v(3), pol(2).v(1), 0, 0, 0, 0, 0, 0, 0)
triA(3) = triangle_number(pol(2).v(0), pol(2).v(3), pol(2).v(2), 0, 0, 0, 0, 0, 0, 0)
 v$ = minus_string(A_ele1.value, A_ele2.value, False, False)
If triangle(triA(0)).data(0).Area <> "" And triangle(triA(1)).data(0).Area = "" Then
   Call add_conditions_to_record(area_of_element_, _
          triangle(triA(0)).data(0).area_no, 0, 0, temp_record.record_data.data0.condition_data)
   combine_two_area_of_element = set_area_of_triangle(triA(1), _
        minus_string(triangle(triA(0)).data(0).Area, v$, True, False), temp_record, 0, 0)
   If combine_two_area_of_element > 1 Then
      Exit Function
   End If
ElseIf triangle(triA(0)).data(0).Area = "" And triangle(triA(1)).data(0).Area <> "" Then
   Call add_conditions_to_record(area_of_element_, _
          triangle(triA(1)).data(0).area_no, 0, 0, temp_record.record_data.data0.condition_data)
   combine_two_area_of_element = set_area_of_triangle(triA(0), _
        add_string(v$, triangle(triA(1)).data(0).Area, True, False), temp_record, 0, 0)
   If combine_two_area_of_element > 1 Then
      Exit Function
   End If
End If
If triangle(triA(2)).data(0).Area <> "" And triangle(triA(3)).data(0).Area = "" Then
   Call add_conditions_to_record(area_of_element_, _
          triangle(triA(2)).data(0).area_no, 0, 0, temp_record.record_data.data0.condition_data)
   combine_two_area_of_element = set_area_of_triangle(triA(3), _
        minus_string(triangle(triA(2)).data(0).Area, v$, True, False), temp_record, 0, 0)
   If combine_two_area_of_element > 1 Then
      Exit Function
   End If
ElseIf triangle(triA(2)).data(0).Area = "" And triangle(triA(3)).data(0).Area <> "" Then
   Call add_conditions_to_record(area_of_element_, _
          triangle(triA(3)).data(0).area_no, 0, 0, temp_record.record_data.data0.condition_data)
   combine_two_area_of_element = set_area_of_triangle(triA(2), _
        add_string(v$, triangle(triA(3)).data(0).Area, True, False), temp_record, 0, 0)
   If combine_two_area_of_element > 1 Then
      Exit Function
   End If
End If
tp% = is_line_line_intersect(line_number0(pol(2).v(0), pol(2).v(2), 0, 0), _
   line_number0(pol(2).v(1), pol(2).v(3), 0, 0), 0, 0, False)
If tp% > 0 Then
triA(0) = triangle_number(pol(2).v(0), pol(2).v(1), tp%, 0, 0, 0, 0, 0, 0, 0)
triA(1) = triangle_number(pol(2).v(3), pol(2).v(2), tp%, 0, 0, 0, 0, 0, 0, 0)
If triangle(triA(0)).data(0).Area <> "" And triangle(triA(1)).data(0).Area = "" Then
   Call add_conditions_to_record(area_of_element_, _
          triangle(triA(0)).data(0).area_no, 0, 0, temp_record.record_data.data0.condition_data)
   combine_two_area_of_element = set_area_of_triangle(triA(1), _
        minus_string(triangle(triA(0)).data(0).Area, v$, True, False), temp_record, 0, 0)
   If combine_two_area_of_element > 1 Then
      Exit Function
   End If
ElseIf triangle(triA(0)).data(0).Area = "" And triangle(triA(1)).data(0).Area <> "" Then
   Call add_conditions_to_record(area_of_element_, _
          triangle(triA(1)).data(0).area_no, 0, 0, temp_record.record_data.data0.condition_data)
   combine_two_area_of_element = set_area_of_triangle(triA(0), _
        add_string(v$, triangle(triA(1)).data(0).Area, True, False), temp_record, 0, 0)
   If combine_two_area_of_element > 1 Then
      Exit Function
   End If
End If
End If
End If
End Function

Public Function combine_two_triangle_with_one_co_point(ByVal l1%, ByVal n10%, ByVal n11%, ByVal n12%, _
       ByVal l2%, ByVal n20%, ByVal n21%, ByVal n22%, po1 As polygon, po2 As polygon) As Integer
Dim dir(1) As Integer
Dim tp(4) As Integer
If n11% > n10% And n12% > n10% Then
   If n11% > n12% Then
    dir(0) = 1
    tp(0) = m_lin(l1%).data(0).data0.in_point(n11%)
    tp(1) = m_lin(l1%).data(0).data0.in_point(n12%)
   Else
    dir(0) = -1
    tp(0) = m_lin(l1%).data(0).data0.in_point(n12%)
    tp(1) = m_lin(l1%).data(0).data0.in_point(n11%)
   End If
ElseIf n11 < n10 And n12 < n10 Then
   If n11 < n12 Then
    dir(0) = 1
    tp(0) = m_lin(l1%).data(0).data0.in_point(n11%)
    tp(1) = m_lin(l1%).data(0).data0.in_point(n12%)
   Else
    dir(0) = -1
    tp(0) = m_lin(l1%).data(0).data0.in_point(n12%)
    tp(1) = m_lin(l1%).data(0).data0.in_point(n11%)
   End If
End If
If n21 > n20 And n22 > n10 Then
   If n21 > n22 Then
    dir(1) = 1
    tp(2) = m_lin(l2%).data(0).data0.in_point(n22%)
    tp(3) = m_lin(l2%).data(0).data0.in_point(n21%)
   Else
    dir(1) = -1
    tp(2) = m_lin(l2%).data(0).data0.in_point(n21%)
    tp(3) = m_lin(l2%).data(0).data0.in_point(n22%)
   End If
ElseIf n21 < n20 And n22 < n20 Then
   If n21 < n22 Then
    dir(1) = 1
    tp(2) = m_lin(l2%).data(0).data0.in_point(n22%)
    tp(3) = m_lin(l2%).data(0).data0.in_point(n21%)
   Else
    dir(1) = -1
    tp(2) = m_lin(l2%).data(0).data0.in_point(n21%)
    tp(3) = m_lin(l2%).data(0).data0.in_point(n22%)
   End If
End If
If dir(0) = 1 And dir(1) = 1 Then
 po1.total_v = 4
 po1.v(0) = tp(0)
 po1.v(1) = tp(1)
 po1.v(2) = tp(2)
 po1.v(3) = tp(3)
 combine_two_triangle_with_one_co_point = -1
ElseIf dir(0) = -1 And dir(1) = -1 Then
 po1.total_v = 4
 po1.v(0) = tp(0)
 po1.v(1) = tp(1)
 po1.v(2) = tp(2)
 po1.v(3) = tp(3)
 combine_two_triangle_with_one_co_point = -2
Else
 po1.total_v = 3
 po1.v(0) = tp(0)
 po1.v(1) = tp(1)
 po1.v(2) = tp(2)
 po1.v(3) = tp(3)
 combine_two_triangle_with_one_co_point = -5
End If
End Function
Public Function combine_v_line_value_with_v_line_value(no%) As Byte
Dim v(2) As Integer
Dim temp_record As total_record_type
Dim eline_data As eline_data0_type
Dim cond_data As condition_data_type
Dim dr_data As relation_data0_type
Dim i%, n%, tl1%, tl2%
Dim re_value$
For i% = 1 To no% - 1
 temp_record.record_data.data0.condition_data.condition_no = 2
 temp_record.record_data.data0.condition_data.condition(1).ty = V_line_value_
 temp_record.record_data.data0.condition_data.condition(1).no = i%
 temp_record.record_data.data0.condition_data.condition(2).ty = V_line_value_
 temp_record.record_data.data0.condition_data.condition(2).no = no%
 If Dtwo_point_line(V_line_value(no%).data(0).v_line).data(0).v_poi(0) = _
        Dtwo_point_line(V_line_value(i%).data(0).v_line).data(0).v_poi(0) Then
       combine_v_line_value_with_v_line_value = _
         combine_two_v_line_value_0(i%, no%, 11)
    If combine_v_line_value_with_v_line_value > 1 Then
     Exit Function
    End If
 ElseIf Dtwo_point_line(V_line_value(no%).data(0).v_line).data(0).v_poi(0) = _
        Dtwo_point_line(V_line_value(i%).data(0).v_line).data(0).v_poi(1) Then
        combine_v_line_value_with_v_line_value = _
         combine_two_v_line_value_0(i%, no%, 21)
    If combine_v_line_value_with_v_line_value > 1 Then
     Exit Function
    End If
 ElseIf Dtwo_point_line(V_line_value(no%).data(0).v_line).data(0).v_poi(1) = _
         Dtwo_point_line(V_line_value(i%).data(0).v_line).data(0).v_poi(0) Then
       combine_v_line_value_with_v_line_value = _
         combine_two_v_line_value_0(i%, no%, 12)
    If combine_v_line_value_with_v_line_value > 1 Then
     Exit Function
    End If
 ElseIf Dtwo_point_line(V_line_value(no%).data(0).v_line).data(0).v_poi(1) = _
         Dtwo_point_line(V_line_value(i%).data(0).v_line).data(0).v_poi(1) Then
        combine_v_line_value_with_v_line_value = _
         combine_two_v_line_value_0(i%, no%, 22)
    If combine_v_line_value_with_v_line_value > 1 Then
     Exit Function
    End If
 ElseIf is_paral_v_line(i%, no%, re_value) Then
    combine_v_line_value_with_v_line_value = set_dparal( _
        Dtwo_point_line(V_line_value(i%).data(0).v_line).data(0).line_no, _
         Dtwo_point_line(V_line_value(no%).data(0).v_line).data(0).line_no, _
          temp_record, 0, 0, False)
     If combine_v_line_value_with_v_line_value > 1 Then
     Exit Function
     End If
  '  combine_v_line_value_with_v_line_value = set_property_of_v_relation( _
         V_line_value(i%).data(0).v_line, _
          V_line_value(no%).data(0).v_line, _
           re_value, temp_record)
  '  If combine_v_line_value_with_v_line_value > 1 Then
  '   Exit Function
  '  End If
  '  If Dtwo_point_line(V_line_value(i%).data(0).v_line).data(0).dir <> _
        Dtwo_point_line(V_line_value(no%).data(0).v_line).data(0).dir Then
  '       re_value = time_string(re_value, "-1", True, False)
  '  End If
  '    combine_v_line_value_with_v_line_value = set_Drelation( _
        Dtwo_point_line(V_line_value(i%).data(0).v_line).data(0).v_poi(0), _
         Dtwo_point_line(V_line_value(i%).data(0).v_line).data(0).v_poi(1), _
          Dtwo_point_line(V_line_value(no%).data(0).v_line).data(0).v_poi(0), _
           Dtwo_point_line(V_line_value(no%).data(0).v_line).data(0).v_poi(1), _
            0, 0, 0, 0, 0, 0, re_value, temp_record, 0, 0, 0, 0, 0)
  '  If combine_v_line_value_with_v_line_value > 1 Then
  '   Exit Function
  '  End If
ElseIf is_verti_v_line(i%, no%, temp_record.record_data.data0.condition_data) Then
    combine_v_line_value_with_v_line_value = set_dverti( _
        Dtwo_point_line(V_line_value(i%).data(0).v_line).data(0).line_no, _
         Dtwo_point_line(V_line_value(no%).data(0).v_line).data(0).line_no, _
          temp_record, 0, 0, False)
    If combine_v_line_value_with_v_line_value > 1 Then
     Exit Function
    End If
 End If
Next i%
End Function

Public Function combine_item0_with_tri_function(ByVal n%) As Byte
Dim i%
If item0(0).data(0).poi(1) < 0 And item0(0).data(0).poi(1) > -5 Then
 For i% = 1 To last_conditions.last_cond(0).tri_function_no
     If tri_function(i%).data(0).A = item0(0).data(0).poi(1) Then
     End If
 Next i%
ElseIf item0(0).data(0).poi(3) < 0 And item0(0).data(0).poi(3) > -5 Then
End If
End Function
Public Function combine_length_of_polygon_with_line_value(ByVal no%) As Byte
Dim i%
Dim t_no%
If conclusion_data(length_of_polygon(no%).record_.conclusion_no - 1).no(0) = 0 Then
 t_no% = last_conditions.last_cond(1).line_value_no
  For i% = 1 To t_no%
         combine_length_of_polygon_with_line_value = _
           combine_line_value_with_length_of_polygon0( _
            i%, no%)
           If combine_length_of_polygon_with_line_value > 1 Then
              Exit Function
           End If
  Next i%
last_combine_length_of_polygon_with_line_value(1) = t_no%
End If
End Function
Public Function combine_length_of_polygon_with_two_line_value(ByVal no%) As Byte
Dim i%
Dim t_no%
If conclusion_data(length_of_polygon(no%).record_.conclusion_no - 1).no(0) = 0 Then
t_no% = last_conditions.last_cond(1).two_line_value_no
For i% = 1 To t_no%
  If two_line_value(i%).data(0).data0.para(0) = "1" And _
      two_line_value(i%).data(0).data0.para(1) = "1" Then
         combine_length_of_polygon_with_two_line_value = _
           combine_two_line_value_with_length_of_polygon0( _
            i%, no%)
           If combine_length_of_polygon_with_two_line_value > 1 Then
              Exit Function
           End If
  End If
Next i%
last_combine_length_of_polygon_with_two_line_value(1) = t_no%
End If
End Function
Public Function combine_length_of_polygon_with_line3_value(ByVal no%) As Byte
Dim i%
Dim t_no%
If conclusion_data(length_of_polygon(no%).record_.conclusion_no - 1).no(0) = 0 Then
t_no% = last_conditions.last_cond(1).line3_value_no
For i% = 1 To t_no%
  If line3_value(i%).data(0).data0.para(0) = "1" And _
   (line3_value(i%).data(0).data0.para(1) = "1" Or _
      line3_value(i%).data(0).data0.para(0) = "-1" Or _
        line3_value(i%).data(0).data0.para(0) = "@1") And _
     (line3_value(i%).data(0).data0.para(0) = "1" Or _
       line3_value(i%).data(0).data0.para(0) = "-1" Or _
        line3_value(i%).data(0).data0.para(0) = "@1") Then
         combine_length_of_polygon_with_line3_value = _
           combine_line3_value_with_length_of_polygon0( _
            i%, no%)
           If combine_length_of_polygon_with_line3_value > 1 Then
              Exit Function
           End If
  End If
Next i%
last_combine_length_of_polygon_with_line3_value(1) = t_no%
End If
End Function

Public Function combine_line_value_with_length_of_polygon(ByVal no%) As Byte
Dim i%
Dim t_no%
t_no% = last_conditions.last_cond(1).length_of_polygon_no
For i% = 1 To t_no%
   If conclusion_data(length_of_polygon(i%).record_.conclusion_no - 1).no(0) = 0 Then
   combine_line_value_with_length_of_polygon = _
     combine_line_value_with_length_of_polygon0(no%, i%)
    If combine_line_value_with_length_of_polygon > 1 Then
       Exit Function
    End If
   End If
Next i%
last_combine_length_of_polygon_with_line_value(0) = t_no%
End Function
Public Function combine_two_line_value_with_length_of_polygon(ByVal no%) As Byte
Dim i%
Dim t_no%
t_no% = last_conditions.last_cond(1).length_of_polygon_no
If two_line_value(no%).data(0).data0.para(0) = "1" And _
    two_line_value(no%).data(0).data0.para(1) = "1" Then
For i% = 1 To t_no%
    If conclusion_data(length_of_polygon(i%).record_.conclusion_no - 1).no(0) = 0 Then
     combine_two_line_value_with_length_of_polygon = _
      combine_two_line_value_with_length_of_polygon0(no%, i%)
    If combine_two_line_value_with_length_of_polygon > 1 Then
       Exit Function
    End If
    End If
 Next i%
 last_combine_length_of_polygon_with_line_value(0) = t_no%
End If
End Function
Public Function combine_line3_value_with_length_of_polygon(ByVal no%) As Byte
Dim i%
Dim t_no%
t_no% = last_conditions.last_cond(1).length_of_polygon_no
If line3_value(no%).data(0).data0.para(0) = "1" And _
   (line3_value(no%).data(0).data0.para(1) = "1" Or _
      line3_value(no%).data(0).data0.para(0) = "-1" Or _
        line3_value(no%).data(0).data0.para(0) = "@1") And _
     (line3_value(no%).data(0).data0.para(0) = "1" Or _
       line3_value(no%).data(0).data0.para(0) = "-1" Or _
        line3_value(no%).data(0).data0.para(0) = "@1") Then
For i% = 1 To t_no%
    If conclusion_data(length_of_polygon(i%).record_.conclusion_no - 1).no(0) = 0 Then
   combine_line3_value_with_length_of_polygon = _
     combine_line3_value_with_length_of_polygon0(no%, i%)
    If combine_line3_value_with_length_of_polygon > 1 Then
       Exit Function
    End If
    End If
 Next i%
last_combine_length_of_polygon_with_line3_value(0) = t_no%
End If
End Function
Public Function combine_line_value_with_length_of_polygon0(ByVal l_v_no%, _
                      ByVal l_of_p_no%) As Byte
Dim l_of_p As length_of_polygon_type
Dim temp_record As total_record_type
Dim para As String
l_of_p = length_of_polygon(l_of_p_no%)
If simple_length_of_polygon_with_segment(l_of_p.data(0), line_value(l_v_no%).data(0).data0.line_no, _
      line_value(l_v_no%).data(0).data0.poi(0), line_value(l_v_no%).data(0).data0.poi(1), para) Then
    l_of_p.data(0).value = add_string(l_of_p.data(0).value, _
         time_string(line_value(l_v_no%).data(0).data0.value, para, False, False), True, False)
    temp_record.record_ = length_of_polygon(l_of_p_no%).record_
    temp_record.record_data.data0.condition_data.condition_no = 2
    temp_record.record_data.data0.condition_data.condition(2).ty = length_of_polygon_
    temp_record.record_data.data0.condition_data.condition(1).ty = line_value_
    temp_record.record_data.data0.condition_data.condition(2).no = l_of_p_no%
    temp_record.record_data.data0.condition_data.condition(1).no = l_v_no%
    temp_record.record_data.data0.theorem_no = 1
    combine_line_value_with_length_of_polygon0 = set_length_of_polygon(l_of_p, 0, temp_record)
End If
End Function
Public Function combine_two_line_value_with_length_of_polygon0(ByVal tl_v_no%, _
                     ByVal l_of_p_no%) As Byte
Dim l_of_p As length_of_polygon_type
Dim temp_record As total_record_type
Dim para(1) As String
l_of_p = length_of_polygon(l_of_p_no%)
If two_line_value(tl_v_no%).data(0).data0.para(0) = "1" And _
               two_line_value(tl_v_no%).data(0).data0.para(1) = "1" Then
   If simple_length_of_polygon_with_segment(l_of_p.data(0), two_line_value(tl_v_no%).data(0).data0.line_no(0), _
      two_line_value(tl_v_no%).data(0).data0.poi(0), two_line_value(tl_v_no%).data(0).data0.poi(1), para(0)) Then
    If simple_length_of_polygon_with_segment(l_of_p.data(0), two_line_value(tl_v_no%).data(0).data0.line_no(1), _
       two_line_value(tl_v_no%).data(0).data0.poi(2), two_line_value(tl_v_no%).data(0).data0.poi(3), para(1)) Then
       If minus_string(para(0), para(1), True, False) = "0" Then
        l_of_p.data(0).value = add_string(l_of_p.data(0).value, time_string( _
            two_line_value(tl_v_no%).data(0).data0.value, para(0), False, False), True, False)
         temp_record.record_ = length_of_polygon(l_of_p_no%).record_
         temp_record.record_data.data0.condition_data.condition_no = 2
         temp_record.record_data.data0.condition_data.condition(2).ty = length_of_polygon_
         temp_record.record_data.data0.condition_data.condition(1).ty = two_line_value_
         temp_record.record_data.data0.condition_data.condition(2).no = l_of_p_no%
         temp_record.record_data.data0.condition_data.condition(1).no = tl_v_no%
         temp_record.record_data.data0.theorem_no = 1
          combine_two_line_value_with_length_of_polygon0 = set_length_of_polygon(l_of_p, 0, temp_record)
       End If
    End If
   End If
End If
End Function
Public Function combine_line3_value_with_length_of_polygon0(ByVal tl_v_no%, _
                           ByVal l_of_p_no%) As Byte
Dim i%, j%
Dim tl(2) As Integer
Dim tp(2, 1) As Integer
Dim tn(1) As Integer
Dim ty As Byte
Dim temp_record As total_record_type
Dim l_of_p As length_of_polygon_type
Dim para(1) As String
l_of_p = length_of_polygon(l_of_p_no%)
If line3_value(tl_v_no%).data(0).data0.value <> "0" Then
  If line3_value(tl_v_no%).data(0).data0.para(0) = "1" And _
    (line3_value(tl_v_no%).data(0).data0.para(1) = "1" Or line3_value(tl_v_no%).data(0).data0.para(1) = "#1") And _
      (line3_value(tl_v_no%).data(0).data0.para(2) = "1" Or line3_value(tl_v_no%).data(0).data0.para(2) = "#1") Then
   ty = 1
  End If
Else
If line3_value(tl_v_no%).data(0).data0.para(0) = "1" And _
    (line3_value(tl_v_no%).data(0).data0.para(1) = "-1" Or line3_value(tl_v_no%).data(0).data0.para(1) = "@1") And _
      (line3_value(tl_v_no%).data(0).data0.para(2) = "-1" Or line3_value(tl_v_no%).data(0).data0.para(2) = "@1") Then
tl(0) = line3_value(tl_v_no%).data(0).data0.line_no(0)
tp(0, 0) = line3_value(tl_v_no%).data(0).data0.poi(0)
tp(0, 1) = line3_value(tl_v_no%).data(0).data0.poi(1)
tn(0) = line3_value(tl_v_no%).data(0).data0.n(0)
tn(1) = line3_value(tl_v_no%).data(0).data0.n(1)
tl(1) = line3_value(tl_v_no%).data(0).data0.line_no(1)
tp(1, 0) = line3_value(tl_v_no%).data(0).data0.poi(2)
tp(1, 1) = line3_value(tl_v_no%).data(0).data0.poi(3)
tl(2) = line3_value(tl_v_no%).data(0).data0.line_no(2)
tp(2, 0) = line3_value(tl_v_no%).data(0).data0.poi(4)
tp(2, 1) = line3_value(tl_v_no%).data(0).data0.poi(5)
ty = 2
ElseIf line3_value(tl_v_no%).data(0).data0.para(0) = "1" And _
    (line3_value(tl_v_no%).data(0).data0.para(1) = "1" Or line3_value(tl_v_no%).data(0).data0.para(1) = "#1") And _
      (line3_value(tl_v_no%).data(0).data0.para(2) = "-1" Or line3_value(tl_v_no%).data(0).data0.para(2) = "@1") Then
tl(0) = line3_value(tl_v_no%).data(0).data0.line_no(2)
tp(0, 0) = line3_value(tl_v_no%).data(0).data0.poi(4)
tp(0, 1) = line3_value(tl_v_no%).data(0).data0.poi(5)
tn(0) = line3_value(tl_v_no%).data(0).data0.n(4)
tn(1) = line3_value(tl_v_no%).data(0).data0.n(5)
tl(1) = line3_value(tl_v_no%).data(0).data0.line_no(1)
tp(1, 0) = line3_value(tl_v_no%).data(0).data0.poi(2)
tp(1, 1) = line3_value(tl_v_no%).data(0).data0.poi(3)
tl(2) = line3_value(tl_v_no%).data(0).data0.line_no(0)
tp(2, 0) = line3_value(tl_v_no%).data(0).data0.poi(0)
tp(2, 1) = line3_value(tl_v_no%).data(0).data0.poi(1)
ty = 2
ElseIf line3_value(tl_v_no%).data(0).data0.para(0) = "1" And _
    (line3_value(tl_v_no%).data(0).data0.para(1) = "-1" Or line3_value(tl_v_no%).data(0).data0.para(1) = "@1") And _
      (line3_value(tl_v_no%).data(0).data0.para(2) = "1" Or line3_value(tl_v_no%).data(0).data0.para(2) = "#1") Then
tl(0) = line3_value(tl_v_no%).data(0).data0.line_no(1)
tp(0, 0) = line3_value(tl_v_no%).data(0).data0.poi(2)
tp(0, 1) = line3_value(tl_v_no%).data(0).data0.poi(3)
tn(0) = line3_value(tl_v_no%).data(0).data0.n(2)
tn(1) = line3_value(tl_v_no%).data(0).data0.n(3)
tl(1) = line3_value(tl_v_no%).data(0).data0.line_no(0)
tp(1, 0) = line3_value(tl_v_no%).data(0).data0.poi(0)
tp(1, 1) = line3_value(tl_v_no%).data(0).data0.poi(1)
tl(2) = line3_value(tl_v_no%).data(0).data0.line_no(2)
tp(2, 0) = line3_value(tl_v_no%).data(0).data0.poi(4)
tp(2, 1) = line3_value(tl_v_no%).data(0).data0.poi(5)
ty = 2
End If
End If
If ty = 1 Then
  If simple_length_of_polygon_with_segment(l_of_p.data(0), _
         line3_value(tl_v_no%).data(0).data0.line_no(0), line3_value(tl_v_no%).data(0).data0.poi(0), _
           line3_value(tl_v_no%).data(0).data0.poi(1), para(0)) Then
   If simple_length_of_polygon_with_segment(l_of_p.data(0), _
          line3_value(tl_v_no%).data(0).data0.line_no(1), line3_value(tl_v_no%).data(0).data0.poi(2), _
           line3_value(tl_v_no%).data(0).data0.poi(3), para(1)) Then
    If minus_string(para(0), para(1), True, False) = "0" Then
     If simple_length_of_polygon_with_segment(l_of_p.data(0), _
         line3_value(tl_v_no%).data(0).data0.line_no(2), line3_value(tl_v_no%).data(0).data0.poi(4), _
           line3_value(tl_v_no%).data(0).data0.poi(5), para(1)) Then
       If minus_string(para(0), para(1), True, False) = "0" Then
        l_of_p.data(0).value = add_string(l_of_p.data(0).value, time_string( _
            line3_value(tl_v_no%).data(0).data0.value, para(0), False, False), True, False)
        temp_record.record_ = length_of_polygon(l_of_p_no%).record_
        temp_record.record_data.data0.condition_data.condition_no = 2
         temp_record.record_data.data0.condition_data.condition(1).ty = length_of_polygon_
          temp_record.record_data.data0.condition_data.condition(2).ty = line3_value_
           temp_record.record_data.data0.condition_data.condition(1).no = l_of_p_no%
            temp_record.record_data.data0.condition_data.condition(2).no = tl_v_no%
             temp_record.record_data.data0.theorem_no = 1
         combine_line3_value_with_length_of_polygon0 = set_length_of_polygon(l_of_p, 0, temp_record)
       End If
     End If
   End If
  End If
 End If
ElseIf ty = 2 Then
     If simple_length_of_polygon_with_segment(l_of_p.data(0), _
         tl(1), tp(1, 0), tp(1, 1), para(0)) Then
      If simple_length_of_polygon_with_segment(l_of_p.data(0), _
          tl(2), tp(2, 0), tp(2, 1), para(1)) Then
        If minus_string(para(0), para(1), True, False) = "0" Then
         If tl(0) < l_of_p.data(0).segment(1).line_no Then
            l_of_p.data(0).last_segment = l_of_p.data(0).last_segment + 1
            For i% = l_of_p.data(0).last_segment To 2 Step -1
              l_of_p.data(0).segment(i%) = l_of_p.data(0).segment(i% - 1)
            Next i%
              l_of_p.data(0).segment(1).line_no = tl(0)
              l_of_p.data(0).segment(1).poi(0) = tp(0, 0)
              l_of_p.data(0).segment(1).poi(1) = tp(0, 1)
              l_of_p.data(0).segment(1).n(0) = tn(0)
              l_of_p.data(0).segment(1).n(0) = tn(1)
              l_of_p.data(0).segment(1).para = para(0)
            GoTo combine_line3_value_with_length_of_polygon0_mark10
         ElseIf tl(0) > l_of_p.data(0).segment(l_of_p.data(0).last_segment).line_no Then
            l_of_p.data(0).last_segment = l_of_p.data(0).last_segment + 1
              l_of_p.data(0).segment(l_of_p.data(0).last_segment).line_no = tl(0)
              l_of_p.data(0).segment(l_of_p.data(0).last_segment).poi(0) = tp(0, 0)
              l_of_p.data(0).segment(l_of_p.data(0).last_segment).poi(1) = tp(0, 1)
              l_of_p.data(0).segment(l_of_p.data(0).last_segment).n(0) = tn(0)
              l_of_p.data(0).segment(l_of_p.data(0).last_segment).n(0) = tn(1)
              l_of_p.data(0).segment(l_of_p.data(0).last_segment).para = para(0)
            GoTo combine_line3_value_with_length_of_polygon0_mark10
         Else
          l_of_p.data(0).segment(l_of_p.data(0).last_segment + 1).line_no = 0
           For i% = 1 To l_of_p.data(0).last_segment
            If tl(0) = l_of_p.data(0).segment(i%).line_no Then
                   If tn(0) < l_of_p.data(0).segment(i%).n(0) And _
                       tn(1) < l_of_p.data(0).segment(i%).n(0) Then
                     For j% = l_of_p.data(0).last_segment To i% + 1 Step -1
                      l_of_p.data(0).segment(j%) = l_of_p.data(0).segment(j% - 1)
                     Next j%
                     l_of_p.data(0).segment(i%).line_no = tl(0)
                     l_of_p.data(0).segment(i%).poi(0) = tp(0, 0)
                     l_of_p.data(0).segment(i%).poi(1) = tp(0, 1)
                     l_of_p.data(0).segment(i%).n(0) = tn(0)
                     l_of_p.data(0).segment(i%).n(0) = tn(1)
                     l_of_p.data(0).segment(i%).para = para(0)
                    GoTo combine_line3_value_with_length_of_polygon0_mark10
                   ElseIf tn(0) < l_of_p.data(0).segment(i%).n(1) And _
                       tn(1) > l_of_p.data(0).segment(i%).n(1) Then
                        Exit Function
                  End If
            ElseIf tl(0) > l_of_p.data(0).segment(i%).line_no And _
                      tl(0) < l_of_p.data(0).segment(i% + 1).line_no Then
              l_of_p.data(0).last_segment = l_of_p.data(0).last_segment + 1
              For j% = l_of_p.data(0).last_segment To i% + 2 Step -1
                l_of_p.data(0).segment(j%) = l_of_p.data(0).segment(j% - 1)
              Next j%
              l_of_p.data(0).segment(i% + 1).line_no = tl(0)
              l_of_p.data(0).segment(i% + 1).poi(0) = tp(0, 0)
              l_of_p.data(0).segment(i% + 1).poi(1) = tp(0, 1)
              l_of_p.data(0).segment(i% + 1).n(0) = tn(0)
              l_of_p.data(0).segment(i% + 1).n(0) = tn(1)
              l_of_p.data(0).segment(i% + 1).para = para(0)
               GoTo combine_line3_value_with_length_of_polygon0_mark10
            End If
          Next i%
         End If
        End If
       End If
     End If
        Exit Function
combine_line3_value_with_length_of_polygon0_mark10:
    l_of_p.data(0).value = add_string(l_of_p.data(0).value, time_string( _
          line3_value(tl_v_no%).data(0).data0.value, para(0), False, False), True, False)
      temp_record.record_ = length_of_polygon(l_of_p_no%).record_
      temp_record.record_data.data0.condition_data.condition_no = 2
      temp_record.record_data.data0.condition_data.condition(2).ty = length_of_polygon_
      temp_record.record_data.data0.condition_data.condition(1).ty = line3_value_
      temp_record.record_data.data0.condition_data.condition(2).no = l_of_p_no%
      temp_record.record_data.data0.condition_data.condition(1).no = tl_v_no%
      temp_record.record_data.data0.theorem_no = 1
      combine_line3_value_with_length_of_polygon0 = set_length_of_polygon(l_of_p, 0, temp_record)
     End If
End Function
Public Function combine_eline_with_length_of_polygon(ByVal el%) As Byte
Dim l_of_p As length_of_polygon_type
Dim para As String
Dim i%, j%, on1%, on2%
Dim temp_record As total_record_type
For i% = 1 To last_conditions.last_cond(1).length_of_polygon_no
If conclusion_data(length_of_polygon(i%).record_.conclusion_no - 1).no(0) = 0 Then
l_of_p = length_of_polygon(i%)
 If simple_length_of_polygon_with_segment(l_of_p.data(0), Deline(el%).data(0).data0.line_no(1), _
      Deline(el%).data(0).data0.poi(2), Deline(el%).data(0).data0.poi(3), para) Then
    If is_line_in_segments(Deline(el%).data(0).data0.line_no(0), Deline(el%).data(0).data0.poi(0), _
          Deline(el%).data(0).data0.poi(1), l_of_p.data(0), on1%, on2%, "") Then
     For j% = on1% To on2%
      l_of_p.data(0).segment(j%).para = add_string(l_of_p.data(0).segment(j%).para, _
            para, True, False)
     Next j%
      temp_record.record_data.data0.condition_data.condition_no = 2
      temp_record.record_data.data0.condition_data.condition(1).ty = eline_
      temp_record.record_data.data0.condition_data.condition(2).ty = length_of_polygon_
      temp_record.record_data.data0.condition_data.condition(1).no = el%
      temp_record.record_data.data0.condition_data.condition(2).no = i%
      temp_record.record_data.data0.theorem_no = 1
      combine_eline_with_length_of_polygon = set_length_of_polygon( _
         l_of_p, 0, temp_record)
         If combine_eline_with_length_of_polygon > 1 Then
           Exit Function
         End If
    End If
 End If
 End If
Next i%
End Function
Public Function simple_length_of_polygon_with_segment(l_of_p As length_of_polygon_type0, _
        ByVal l%, ByVal p1%, ByVal p2%, para As String) As Boolean
Dim i%, j%, k%
If is_line_in_segments(l%, p1%, p2%, l_of_p, i%, k%, para) Then
          For j% = k% To l_of_p.last_segment - 1
            l_of_p.segment(i% + (j% - k%)) = l_of_p.segment(j% + 1)
          Next j%
                    l_of_p.last_segment = l_of_p.last_segment + i% - k% - 1
                     simple_length_of_polygon_with_segment = True
End If
End Function

Public Function is_line_in_segments(ByVal l%, ByVal p1%, ByVal p2%, _
         l_of_p As length_of_polygon_type0, o_n1%, o_n2%, para As String) As Boolean
Dim i%, k%
For i% = 1 To l_of_p.last_segment
 If l_of_p.segment(i%).line_no = l% Then
  If l_of_p.segment(i%).poi(0) = p1% Then
     o_n1% = i%
   para = l_of_p.segment(i%).para
            k% = i%
        Do While k% <= l_of_p.last_segment
         If l_of_p.segment(k%).line_no <> l% Or l_of_p.segment(k%).para <> para Then
            Exit Function
         ElseIf l_of_p.segment(k%).poi(1) = p2% Then
            is_line_in_segments = True
             o_n2% = k%
            Exit Function
         End If
         k% = k% + 1
        Loop
         Exit Function
  End If
 End If
Next i%
End Function
Public Function combine_angle3_value_with_item0(ByVal A3_v_n%, ByVal k1%, ByVal it_n%, ByVal k2%) As Byte
Dim tA(1) As Integer
Dim ty As Integer
Dim tp(3) As Integer
Dim para As String
Dim sig As String
Dim temp_record0_ As condition_data_type
sig = item0(it_n%).data(0).sig
tA(0) = angle3_value(A3_v_n%).data(0).data0.angle(k1%)
tA(1) = angle3_value(A3_v_n%).data(0).data0.angle((k1% + 1) Mod 2)
If angle3_value(A3_v_n%).data(0).data0.value = "0" Then
   If angle3_value(A3_v_n%).data(0).data0.para(0) = "1" And _
       angle3_value(A3_v_n%).data(0).data0.para(1) = "-1" Then
       ty = 1
   ElseIf angle3_value(A3_v_n%).data(0).data0.para(0) = "1" And _
       angle3_value(A3_v_n%).data(0).data0.para(1) = "-2" Then
        If k1% = 0 Then
           ty = 6
        End If
   ElseIf angle3_value(A3_v_n%).data(0).data0.para(0) = "2" And _
       angle3_value(A3_v_n%).data(0).data0.para(1) = "-1" Then
        If k1% = 1 Then
           ty = 6
        End If
   End If
ElseIf angle3_value(A3_v_n%).data(0).data0.value = "90" Then
   If angle3_value(A3_v_n%).data(0).data0.para(0) = "1" And _
       angle3_value(A3_v_n%).data(0).data0.para(1) = "1" Then
    ty = 2
   ElseIf angle3_value(A3_v_n%).data(0).data0.para(0) = "1" And _
      angle3_value(A3_v_n%).data(0).data0.para(1) = "-1" Then
    If k1% = 0 Then
     ty = 3
    Else
     ty = 4
    End If
   End If
ElseIf angle3_value(A3_v_n%).data(0).data0.value = "-90" Or _
        angle3_value(A3_v_n%).data(0).data0.value = "@90" Then
   ElseIf angle3_value(A3_v_n%).data(0).data0.para(0) = "1" And _
             angle3_value(A3_v_n%).data(0).data0.para(1) = "-1" Then
    If k1% = 0 Then
     ty = 4
    Else
     ty = 3
   End If
ElseIf angle3_value(A3_v_n%).data(0).data0.value = "180" Then
   If angle3_value(A3_v_n%).data(0).data0.para(0) = "1" And _
     angle3_value(A3_v_n%).data(0).data0.para(1) = "1" Then
    ty = 5
   End If
End If
If ty = 0 Then
    Exit Function
End If
tp(0) = item0(it_n%).data(0).poi(2 * k2%)
tp(1) = item0(it_n%).data(0).poi(2 * k2% + 1)
tp(2) = item0(it_n%).data(0).poi(2 * (k2% + 1) Mod 2)
tp(3) = item0(it_n%).data(0).poi((2 * (k2% + 1) Mod 2) + 1)
If tp(1) < 0 And tp(1) > -5 Then
   para = "1"
   If ty = 1 Then
    tp(0) = tA(1)
   ElseIf ty = 2 Or ty = 3 Or ty = 4 Then
   If ty = 3 Then
    If tp(1) > -1 Then
       para = "-1"
    End If
   ElseIf ty = 4 Then
     If tp(1) = -1 Or tp(1) = -3 Or tp(1) = -4 Then
      para = "-1"
     End If
   End If
    If tp(1) = -1 Then
       tp(1) = -2
    ElseIf tp(1) = -2 Then
       tp(1) = -1
    ElseIf tp(1) = -3 Then
       tp(1) = -4
    ElseIf tp(1) = -4 Then
       tp(1) = -3
    End If
       tp(0) = tA(1)
   ElseIf ty = 5 Then
    If tp(1) < -1 Then
     para = "-"
    End If
   ElseIf ty = 6 Then
    If tp(0) = tp(2) Then
       If (tp(1) = -1 And tp(3) = -2) Or (tp(1) = -2 And tp(1) = -1) Then
        para = "1/2"
         tp(0) = tA(1)
         tp(1) = -1
         tp(2) = 0
         tp(3) = 0
         sig = "~"
         '倍角公式
       Else
        Exit Function
       End If
    Else
     Exit Function
    End If
   End If
 If k2% = 1 Then
    Call exchange_two_integer(tp(0), tp(2))
    Call exchange_two_integer(tp(1), tp(3))
 End If
 If sig = "/" Then
    If tp(0) = tp(2) And tp(1) = tp(3) Then
       tp(0) = 0
       tp(1) = 0
       tp(2) = 0
       tp(3) = 0
    End If
 ElseIf sig = "*" Then
       If tp(0) = tp(2) Then
          If (tp(1) = -3 And tp(3) = -4) Or (tp(1) = -4 And tp(3) = -3) Then
           tp(0) = 0
           tp(1) = 0
           tp(2) = 0
           tp(3) = 0
          End If
       End If
 End If
  temp_record0_.condition_no = 0
  Call add_conditions_to_record(angle3_value_, A3_v_n%, 0, 0, temp_record0_)
  Call add_conditions_to_record(item0_, it_n%, 0, 0, temp_record0_)
   combine_angle3_value_with_item0 = set_item0(tp(0), tp(1), tp(2), tp(3), sig, _
        0, 0, 0, 0, 0, 0, "1", "1", para, "", "1", 0, _
         temp_record0_, 0, it_n%, 0, 0, condition_data0, False)
End If
End Function
Public Function combine_angle3_value_with_item1(ByVal A3_v_n%, ByVal k1%, ByVal it_n%, ByVal k2%) As Byte
Dim no%
Dim tv$
Dim temp_record As total_record_type
record_0.data0.condition_data.condition_no = 0
  If item0(it_n%).data(0).poi(2 * (k1% + 1) Mod 2) <> 0 Then
   Call set_item0(item0(it_n%).data(0).poi(2 * (k1% + 1) Mod 2), _
    item0(it_n%).data(0).poi(2 * (k1% + 1) Mod 2 + 1), 0, 0, "~", 0, 0, _
     0, 0, 0, 0, "1", "1", "1", "", "1", item0(it_n%).data(0).conclusion_no, _
       record_0.data0.condition_data, -1, no%, 0, 0, record_0.data0.condition_data, False)
  Else
    no% = 0
  End If
  If item0(it_n%).data(0).poi(2 * k1% + 1) = -1 Then
   tv$ = sin_(angle3_value(A3_v_n%).data(0).data0.value, 0)
  ElseIf item0(it_n%).data(0).poi(2 * k1% + 1) = -2 Then
   tv$ = cos_(angle3_value(A3_v_n%).data(0).data0.value, 0)
  ElseIf item0(it_n%).data(0).poi(2 * k1% + 1) = -3 Then
   tv$ = tan_(angle3_value(A3_v_n%).data(0).data0.value, 0)
  ElseIf item0(it_n%).data(0).poi(2 * k1% + 1) = -4 Then
   tv$ = tan_(angle3_value(A3_v_n%).data(0).data0.value, 0)
   tv$ = divide_string("1", tv$, True, False)
  End If
Call add_conditions_to_record(angle3_value_, A3_v_n%, 0, 0, temp_record.record_data.data0.condition_data)
If tv$ <> "F" Then
If no% > 0 Then
Call add_conditions_to_record(item0_, it_n%, 0, 0, temp_record.record_data.data0.condition_data)
 'combine_angle3_value_with_item1=set_item0_value(no%,
Else
 combine_angle3_value_with_item1 = set_item0_value(it_n%, 0, 0, "0", "0", _
              tv$, "", 0, temp_record.record_data.data0.condition_data)
End If
End If
End Function
Public Function combine_angle3_value_with_item(ByVal A3_n%) As Byte
Dim i%, j%, k%, l%, no%
Dim n_(1) As Integer
Dim tn() As Integer
Dim ite As item0_data_type
If angle3_value(A3_n%).data(0).data0.para(2) = "0" Then
 If angle3_value(A3_n%).data(0).data0.para(1) = "0" Then
      For i% = 0 To 1
       For j% = 0 To 1
        ite.poi(2 * j%) = angle3_value(A3_n%).data(0).data0.angle(i%)
        ite.poi(2 * j% + 1) = -5
         Call search_for_item0(ite, j%, n_(0), 1)
        ite.poi(2 * j% + 1) = 0
         Call search_for_item0(ite, j%, n_(1), 1)
        For l% = n_(0) + 1 To n_(1)
         no% = item0(l%).data(0).index(i%)
          If item0(no%).data(0).value = "" Then
          ReDim Preserve tn(k%) As Integer
           tn(k%) = no%
            k% = k% + 1
          End If
        Next l%
        For l% = 0 To k% - 1
         no% = tn(l%)
          combine_angle3_value_with_item = combine_angle3_value_with_item1( _
               A3_n%, i%, no%, j%)
           If combine_angle3_value_with_item > 1 Then
              Exit Function
           End If
        Next l%
       Next j%
      Next i%
  Else
 If angle3_value(A3_n%).data(0).data0.value = "90" Or _
     angle3_value(A3_n%).data(0).data0.value = "@90" Or _
      angle3_value(A3_n%).data(0).data0.value = "0" Or _
       angle3_value(A3_n%).data(0).data0.value = "180" Then
    If angle3_value(A3_n%).data(0).data0.para(0) = "1" Or _
         angle3_value(A3_n%).data(0).data0.para(0) = "2" Then
     If angle3_value(A3_n%).data(0).data0.para(0) = "1" Or _
         angle3_value(A3_n%).data(0).data0.para(0) = "2" Or _
           angle3_value(A3_n%).data(0).data0.para(0) = "-1" Or _
            angle3_value(A3_n%).data(0).data0.para(0) = "-2" Or _
             angle3_value(A3_n%).data(0).data0.para(0) = "@1" Or _
              angle3_value(A3_n%).data(0).data0.para(0) = "@2" Then
      For i% = 0 To 1
       For j% = 0 To 1
        ite.poi(2 * j%) = angle3_value(A3_n%).data(0).data0.angle(i%)
        ite.poi(2 * j% + 1) = -5
         Call search_for_item0(ite, j%, n_(0), 1)
        ite.poi(2 * j% + 1) = 0
         Call search_for_item0(ite, j%, n_(1), 1)
        For l% = n_(0) + 1 To n_(1)
         no% = item0(l%).data(0).index(i%)
          ReDim Preserve tn(k%) As Integer
           tn(k%) = no%
            k% = k% + 1
        Next l%
        For l% = 0 To k% - 1
         no% = tn(l%)
          combine_angle3_value_with_item = combine_angle3_value_with_item0( _
               A3_n%, i%, no%, j%)
           If combine_angle3_value_with_item > 1 Then
              Exit Function
           End If
        Next l%
       Next j%
      Next i%
     End If
    End If
 End If
End If
End If
End Function
Public Function combine_item_with_angle3_value(ByVal it_n%) As Byte
Dim i%, j%, k%, l%, no%
Dim n_(1) As Integer
Dim tn() As Integer
Dim A3_v As angle3_value_data0_type
 If item0(it_n%).data(0).poi(1) < 0 Or item0(it_n%).data(0).poi(3) < 0 Then
     For i% = 0 To 1
       For j% = 0 To 1
        If item0(it_n%).data(0).poi(2 * j%) > 0 Then
         A3_v.angle(i%) = item0(it_n%).data(0).poi(2 * j%)
         A3_v.angle((i% + 2) Mod 3) = -1
          Call search_for_three_angle_value(A3_v, j%, n_(0), 1)   '5.7
         A3_v.angle((i% + 2) Mod 3) = 30000
          Call search_for_three_angle_value(A3_v, j%, n_(1), 1)   '5.7
For l% = n_(0) + 1 To n_(1)
no% = angle3_value(l%).data(0).record.data1.index.i(i%)
If angle3_value(no%).data(0).data0.para(2) = "0" And _
     angle3_value(no%).data(0).data0.para(1) <> "0" Then
 If angle3_value(no%).data(0).data0.value = "90" Or _
     angle3_value(no%).data(0).data0.value = "@90" Or _
      angle3_value(no%).data(0).data0.value = "0" Or _
       angle3_value(no%).data(0).data0.value = "180" Then
    If angle3_value(no%).data(0).data0.para(0) = "1" Or _
         angle3_value(no%).data(0).data0.para(0) = "2" Then
     If angle3_value(no%).data(0).data0.para(0) = "1" Or _
         angle3_value(no%).data(0).data0.para(0) = "2" Or _
           angle3_value(no%).data(0).data0.para(0) = "-1" Or _
            angle3_value(no%).data(0).data0.para(0) = "-2" Or _
             angle3_value(no%).data(0).data0.para(0) = "@1" Or _
              angle3_value(no%).data(0).data0.para(0) = "@2" Then
        ReDim Preserve tn(k%) As Integer
           tn(k%) = no%
            k% = k% + 1
     End If
    End If
  End If
End If
Next l%
        For l% = 0 To k% - 1
         no% = tn(l%)
          combine_item_with_angle3_value = combine_angle3_value_with_item0( _
               no%, i%, it_n%, j%)
           If combine_item_with_angle3_value > 1 Then
              Exit Function
           End If
        Next l%
     End If
    Next j%
   Next i%
End If
End Function
Public Function combine_general_string_with_two_line_value(ByVal ge%, ByVal tlv%) As Byte
If tlv% > 0 Then
  If two_line_value(tlv%).data(0).data0.para(0) = "1" And _
       (two_line_value(tlv%).data(0).data0.para(1) = "1" Or _
         two_line_value(tlv%).data(0).data0.para(0) = "-1") Then
         For ge% = 1 To last_conditions.last_cond(1).general_string_no
             If general_string(ge%).data(0).para(2) = "0" Then
                If general_string(ge%).data(0).para(0) = "1" And _
                     (general_string(ge%).data(0).para(1) = "1" Or _
                        general_string(ge%).data(0).para(1) = "-1") Then
                   If item0(general_string(ge%).data(0).item(0)).data(0).sig = "*" And _
                        item0(general_string(ge%).data(0).item(1)).data(0).sig = "*" Then
                   If item0(general_string(ge%).data(0).item(0)).data(0).poi(0) = _
                       item0(general_string(ge%).data(0).item(0)).data(0).poi(2) And _
                      item0(general_string(ge%).data(0).item(0)).data(0).poi(1) = _
                       item0(general_string(ge%).data(0).item(0)).data(0).poi(3) And _
                      item0(general_string(ge%).data(0).item(1)).data(0).poi(0) = _
                       item0(general_string(ge%).data(0).item(1)).data(0).poi(2) And _
                      item0(general_string(ge%).data(0).item(1)).data(0).poi(1) = _
                       item0(general_string(ge%).data(0).item(1)).data(0).poi(3) Then
                      If (item0(general_string(ge%).data(0).item(0)).data(0).poi(0) = _
                            two_line_value(tlv%).data(0).data0.poi(0) And _
                           item0(general_string(ge%).data(0).item(0)).data(0).poi(1) = _
                            two_line_value(tlv%).data(0).data0.poi(1) And _
                           item0(general_string(ge%).data(0).item(1)).data(0).poi(0) = _
                            two_line_value(tlv%).data(0).data0.poi(2) And _
                           item0(general_string(ge%).data(0).item(1)).data(0).poi(1) = _
                            two_line_value(tlv%).data(0).data0.poi(3)) Or _
                          (item0(general_string(ge%).data(0).item(1)).data(0).poi(0) = _
                            two_line_value(tlv%).data(0).data0.poi(0) And _
                           item0(general_string(ge%).data(0).item(1)).data(0).poi(1) = _
                            two_line_value(tlv%).data(0).data0.poi(1) And _
                           item0(general_string(ge%).data(0).item(0)).data(0).poi(0) = _
                            two_line_value(tlv%).data(0).data0.poi(2) And _
                            item0(general_string(ge%).data(0).item(0)).data(0).poi(1) = _
                            two_line_value(tlv%).data(0).data0.poi(3)) Then
                        combine_general_string_with_two_line_value = _
                         combine_general_string_with_two_line_value0(ge%, tlv%)
                          If combine_general_string_with_two_line_value > 1 Then
                             Exit Function
                          End If
                        End If
                    End If
                   End If
                End If
             End If
         Next ge%
  End If
Else
 If general_string(ge%).data(0).para(2) = "0" Then
    If general_string(ge%).data(0).para(0) = "1" And _
         (general_string(ge%).data(0).para(1) = "1" Or _
             general_string(ge%).data(0).para(1) = "-1") Then
          If item0(general_string(ge%).data(0).item(0)).data(0).sig = "*" And _
               item0(general_string(ge%).data(0).item(1)).data(0).sig = "*" Then
                If item0(general_string(ge%).data(0).item(0)).data(0).poi(0) = _
                   item0(general_string(ge%).data(0).item(0)).data(0).poi(2) And _
                   item0(general_string(ge%).data(0).item(0)).data(0).poi(1) = _
                   item0(general_string(ge%).data(0).item(0)).data(0).poi(3) And _
                   item0(general_string(ge%).data(0).item(1)).data(0).poi(0) = _
                   item0(general_string(ge%).data(0).item(1)).data(0).poi(2) And _
                   item0(general_string(ge%).data(0).item(1)).data(0).poi(1) = _
                   item0(general_string(ge%).data(0).item(1)).data(0).poi(3) Then
      For tlv% = 1 To last_conditions.last_cond(1).two_line_value_no
        If two_line_value(tlv%).data(0).data0.para(0) = "1" And _
           (two_line_value(tlv%).data(0).data0.para(1) = "1" Or _
             two_line_value(tlv%).data(0).data0.para(0) = "-1") Then
                      If (item0(general_string(ge%).data(0).item(0)).data(0).poi(0) = _
                            two_line_value(tlv%).data(0).data0.poi(0) And _
                           item0(general_string(ge%).data(0).item(0)).data(0).poi(1) = _
                            two_line_value(tlv%).data(0).data0.poi(1) And _
                           item0(general_string(ge%).data(0).item(1)).data(0).poi(0) = _
                            two_line_value(tlv%).data(0).data0.poi(2) And _
                           item0(general_string(ge%).data(0).item(1)).data(0).poi(1) = _
                            two_line_value(tlv%).data(0).data0.poi(3)) Or _
                          (item0(general_string(ge%).data(0).item(1)).data(0).poi(0) = _
                            two_line_value(tlv%).data(0).data0.poi(0) And _
                           item0(general_string(ge%).data(0).item(1)).data(0).poi(1) = _
                            two_line_value(tlv%).data(0).data0.poi(1) And _
                           item0(general_string(ge%).data(0).item(0)).data(0).poi(0) = _
                            two_line_value(tlv%).data(0).data0.poi(2) And _
                            item0(general_string(ge%).data(0).item(0)).data(0).poi(1) = _
                            two_line_value(tlv%).data(0).data0.poi(3)) Then
                        combine_general_string_with_two_line_value = _
                         combine_general_string_with_two_line_value0(ge%, tlv%)
                          If combine_general_string_with_two_line_value > 1 Then
                             Exit Function
                          End If
                        End If
         End If
      Next tlv%
      End If
     End If
    End If
  End If
End If
End Function
Public Function combine_general_string_with_two_line_value0(ByVal ge%, ByVal tlv%) As Byte
Dim temp_record As total_record_type
Dim v$
Dim tv(1) As String
Call add_conditions_to_record(general_string_, ge%, 0, 0, temp_record.record_data.data0.condition_data)
Call add_conditions_to_record(two_line_value_, tlv%, 0, 0, temp_record.record_data.data0.condition_data)
temp_record.record_data.data0.theorem_no = 1
If general_string(ge%).data(0).para(1) = "-1" Then
v$ = divide_string(general_string(ge%).data(0).value, two_line_value(tlv%).data(0).data0.value, _
                        True, False)
tv(0) = add_string(v$, two_line_value(tlv%).data(0).data0.value, True, False)
tv(0) = divide_string(tv(0), "2", True, False)
tv(1) = minus_string(two_line_value(tlv%).data(0).data0.value, v$, True, False)
tv(1) = divide_string(tv(1), "2", True, False)
If two_line_value(tlv%).data(0).data0.para(1) = "-1" Then
tv(1) = time_string(tv(1), "-1", True, False)
End If
combine_general_string_with_two_line_value0 = _
     set_line_value(item0(general_string(ge%).data(0).item(0)).data(0).poi(0), _
         item0(general_string(ge%).data(0).item(0)).data(0).poi(1), _
          tv(0), item0(general_string(ge%).data(0).item(0)).data(0).n(0), _
           item0(general_string(ge%).data(0).item(0)).data(0).n(1), _
            item0(general_string(ge%).data(0).item(0)).data(0).line_no(0), temp_record, 0, 0, False)
   If combine_general_string_with_two_line_value0 > 1 Then
    Exit Function
   End If
combine_general_string_with_two_line_value0 = _
     set_line_value(item0(general_string(ge%).data(0).item(1)).data(0).poi(0), _
         item0(general_string(ge%).data(0).item(1)).data(0).poi(1), _
          tv(1), item0(general_string(ge%).data(0).item(1)).data(0).n(0), _
           item0(general_string(ge%).data(0).item(1)).data(0).n(1), _
            item0(general_string(ge%).data(0).item(1)).data(0).line_no(0), temp_record, 0, 0, False)
   If combine_general_string_with_two_line_value0 > 1 Then
    Exit Function
   End If
 Else
  v$ = time_string(two_line_value(tlv%).data(0).data0.value, _
                   two_line_value(tlv%).data(0).data0.value, True, False)
  v$ = minus_string(general_string(ge%).data(0).value, v$, True, False)
     v$ = add_string(general_string(ge%).data(0).value, v$, True, False)
      If val(v$) < 0 Then
        error_of_wenti = 1
         combine_general_string_with_two_line_value0 = 2
          Exit Function
      Else
       v$ = sqr_string(v$, True, False)
      End If
     tv(0) = add_string(v$, two_line_value(tlv%).data(0).data0.value, False, False)
     tv(0) = divide_string(tv(0), "2", True, False)
     If two_line_value(tlv%).data(0).data0.para(0) = "1" Then
      tv(1) = minus_string(two_line_value(tlv%).data(0).data0.value, v$, False, False)
     Else
      tv(1) = minus_string(v$, two_line_value(tlv%).data(0).data0.value, False, False)
     End If
     tv(1) = divide_string(tv(1), "2", True, False)
     If (m_poi(two_line_value(tlv%).data(0).data0.poi(0)).data(0).data0.coordinate.X - _
         m_poi(two_line_value(tlv%).data(0).data0.poi(1)).data(0).data0.coordinate.X) ^ 2 + _
         (m_poi(two_line_value(tlv%).data(0).data0.poi(0)).data(0).data0.coordinate.Y - _
          m_poi(two_line_value(tlv%).data(0).data0.poi(1)).data(0).data0.coordinate.Y) ^ 2 > _
          (m_poi(two_line_value(tlv%).data(0).data0.poi(2)).data(0).data0.coordinate.X - _
         m_poi(two_line_value(tlv%).data(0).data0.poi(3)).data(0).data0.coordinate.X) ^ 2 + _
          (m_poi(two_line_value(tlv%).data(0).data0.poi(2)).data(0).data0.coordinate.Y - _
         m_poi(two_line_value(tlv%).data(0).data0.poi(3)).data(0).data0.coordinate.Y) ^ 2 Then
        combine_general_string_with_two_line_value0 = _
         set_line_value(two_line_value(tlv%).data(0).data0.poi(0), _
           two_line_value(tlv%).data(0).data0.poi(1), _
             tv(0), two_line_value(tlv%).data(0).data0.n(0), _
               two_line_value(tlv%).data(0).data0.n(1), _
                 two_line_value(tlv%).data(0).data0.line_no(0), temp_record, 0, 0, False)
          If combine_general_string_with_two_line_value0 > 1 Then
           Exit Function
          End If
         combine_general_string_with_two_line_value0 = _
          set_line_value(two_line_value(tlv%).data(0).data0.poi(2), _
            two_line_value(tlv%).data(0).data0.poi(3), _
              tv(1), two_line_value(tlv%).data(0).data0.n(2), _
               two_line_value(tlv%).data(0).data0.n(3), _
                two_line_value(tlv%).data(0).data0.line_no(1), temp_record, 0, 0, False)
          If combine_general_string_with_two_line_value0 > 1 Then
           Exit Function
          End If
      Else
        combine_general_string_with_two_line_value0 = _
         set_line_value(two_line_value(tlv%).data(0).data0.poi(0), _
           two_line_value(tlv%).data(0).data0.poi(1), _
             tv(1), two_line_value(tlv%).data(0).data0.n(0), _
               two_line_value(tlv%).data(0).data0.n(1), _
                 two_line_value(tlv%).data(0).data0.line_no(0), temp_record, 0, 0, False)
          If combine_general_string_with_two_line_value0 > 1 Then
           Exit Function
          End If
         combine_general_string_with_two_line_value0 = _
          set_line_value(two_line_value(tlv%).data(0).data0.poi(2), _
            two_line_value(tlv%).data(0).data0.poi(3), _
              tv(0), two_line_value(tlv%).data(0).data0.n(2), _
               two_line_value(tlv%).data(0).data0.n(3), _
                two_line_value(tlv%).data(0).data0.line_no(1), temp_record, 0, 0, False)
          If combine_general_string_with_two_line_value0 > 1 Then
           Exit Function
          End If
      End If
 End If
End Function


Public Function combine_two_equal_arc(ByVal no%) As Byte
Dim i%, j%, tA%
Dim temp_record As total_record_type
For i% = 1 To last_conditions.last_cond(1).equal_arc_no
 If i% <> no% Then
 temp_record.record_data.data0.condition_data.condition_no = 0
 Call add_conditions_to_record(equal_arc_, no%, i%, 0, temp_record.record_data.data0.condition_data)
 temp_record.record_data.data0.theorem_no = 1
 If equal_arc(i%).data(0).arc(0) = equal_arc(no%).data(0).arc(0) Then
 combine_two_equal_arc = set_equal_arc(equal_arc(i%).data(0).arc(1), _
            equal_arc(no%).data(0).arc(1), temp_record, 0, 0)
   If combine_two_equal_arc > 1 Then
      Exit Function
   End If
 ElseIf equal_arc(i%).data(0).arc(0) = equal_arc(no%).data(0).arc(1) Then
 combine_two_equal_arc = set_equal_arc(equal_arc(i%).data(0).arc(1), _
            equal_arc(no%).data(0).arc(0), temp_record, 0, 0)
   If combine_two_equal_arc > 1 Then
      Exit Function
   End If
 ElseIf equal_arc(i%).data(0).arc(1) = equal_arc(no%).data(0).arc(0) Then
 combine_two_equal_arc = set_equal_arc(equal_arc(i%).data(0).arc(0), _
            equal_arc(no%).data(0).arc(1), temp_record, 0, 0)
   If combine_two_equal_arc > 1 Then
      Exit Function
   End If
 ElseIf equal_arc(i%).data(0).arc(1) = equal_arc(no%).data(0).arc(1) Then
 combine_two_equal_arc = set_equal_arc(equal_arc(i%).data(0).arc(0), _
            equal_arc(no%).data(0).arc(0), temp_record, 0, 0)
   If combine_two_equal_arc > 1 Then
      Exit Function
   End If
 ElseIf equal_arc(i%).data(0).arc(2) = equal_arc(no%).data(0).arc(2) And _
          equal_arc(i%).data(0).arc(2) > 0 Then
  combine_two_equal_arc = set_equal_arc(equal_arc(i%).data(0).arc(0), _
            equal_arc(no%).data(0).arc(0), temp_record, 0, 0)
   If combine_two_equal_arc > 1 Then
      Exit Function
   End If
  combine_two_equal_arc = set_equal_arc(equal_arc(i%).data(0).arc(0), _
            equal_arc(no%).data(0).arc(1), temp_record, 0, 0)
   If combine_two_equal_arc > 1 Then
      Exit Function
   End If
  combine_two_equal_arc = set_equal_arc(equal_arc(i%).data(0).arc(1), _
            equal_arc(no%).data(0).arc(0), temp_record, 0, 0)
   If combine_two_equal_arc > 1 Then
      Exit Function
   End If
  combine_two_equal_arc = set_equal_arc(equal_arc(i%).data(0).arc(1), _
            equal_arc(no%).data(0).arc(1), temp_record, 0, 0)
   If combine_two_equal_arc > 1 Then
      Exit Function
   End If
 Else
    If equal_arc(no%).data(0).arc(0) = equal_arc(i%).data(0).arc(2) Then
      tA% = equal_arc(no%).data(0).arc(1)
    ElseIf equal_arc(no%).data(0).arc(1) = equal_arc(i%).data(0).arc(2) Then
      tA% = equal_arc(no%).data(0).arc(0)
    Else
     tA% = 0
    End If
    If tA% > 0 Then
      For j% = i% + 1 To last_conditions.last_cond(1).equal_arc_no
        If j% <> no% Then
          If tA% = equal_arc(j%).data(0).arc(2) Then
            temp_record.record_data.data0.condition_data.condition_no = 0
            Call add_conditions_to_record(equal_arc_, no%, i%, j%, temp_record.record_data.data0.condition_data)
   combine_two_equal_arc = set_equal_arc(equal_arc(i%).data(0).arc(0), _
            equal_arc(j%).data(0).arc(0), temp_record, 0, 0)
   If combine_two_equal_arc > 1 Then
      Exit Function
   End If
  combine_two_equal_arc = set_equal_arc(equal_arc(i%).data(0).arc(0), _
            equal_arc(j%).data(0).arc(1), temp_record, 0, 0)
   If combine_two_equal_arc > 1 Then
      Exit Function
   End If
  combine_two_equal_arc = set_equal_arc(equal_arc(i%).data(0).arc(1), _
            equal_arc(j%).data(0).arc(0), temp_record, 0, 0)
   If combine_two_equal_arc > 1 Then
      Exit Function
   End If
  combine_two_equal_arc = set_equal_arc(equal_arc(i%).data(0).arc(1), _
            equal_arc(j%).data(0).arc(1), temp_record, 0, 0)
   If combine_two_equal_arc > 1 Then
      Exit Function
   End If
    End If
    End If
      Next j%
    End If
 End If
 End If
Next i%
End Function

Public Function combine_equal_arc_with_arc_value(ByVal e_ac%, ByVal ac_v%) As Byte
Dim temp_record As total_record_type
If e_ac% > 0 Then
  For ac_v% = 1 To last_conditions.last_cond(1).arc_value_no
  temp_record.record_data.data0.condition_data.condition_no = 0
  Call add_conditions_to_record(arc_value_, ac_v%, 0, 0, temp_record.record_data.data0.condition_data)
  Call add_conditions_to_record(equal_arc_, e_ac%, 0, 0, temp_record.record_data.data0.condition_data)
 If equal_arc(e_ac%).data(0).arc(0) = arc_value(ac_v%).data(0).arc Then
 combine_equal_arc_with_arc_value = set_arc_value(equal_arc(e_ac%).data(0).arc(1), _
          arc_value(ac_v%).data(0).value, temp_record, 0, 0)
   If combine_equal_arc_with_arc_value > 1 Then
      Exit Function
   End If
 ElseIf equal_arc(e_ac%).data(0).arc(1) = arc_value(ac_v%).data(0).arc Then
 combine_equal_arc_with_arc_value = set_arc_value(equal_arc(e_ac%).data(0).arc(0), _
          arc_value(ac_v%).data(0).value, temp_record, 0, 0)
   If combine_equal_arc_with_arc_value > 1 Then
      Exit Function
   End If
 End If
Next ac_v%
Else
For e_ac% = 1 To last_conditions.last_cond(1).equal_arc_no
  temp_record.record_data.data0.condition_data.condition_no = 0
  Call add_conditions_to_record(arc_value_, ac_v%, 0, 0, temp_record.record_data.data0.condition_data)
  Call add_conditions_to_record(equal_arc_, e_ac%, 0, 0, temp_record.record_data.data0.condition_data)
 If equal_arc(e_ac%).data(0).arc(0) = arc_value(ac_v%).data(0).arc Then
 combine_equal_arc_with_arc_value = set_arc_value(equal_arc(e_ac%).data(0).arc(1), _
          arc_value(ac_v%).data(0).value, temp_record, 0, 0)
   If combine_equal_arc_with_arc_value > 1 Then
      Exit Function
   End If
 ElseIf equal_arc(e_ac%).data(0).arc(1) = arc_value(ac_v%).data(0).arc Then
 combine_equal_arc_with_arc_value = set_arc_value(equal_arc(e_ac%).data(0).arc(0), _
          arc_value(ac_v%).data(0).value, temp_record, 0, 0)
   If combine_equal_arc_with_arc_value > 1 Then
      Exit Function
   End If
 End If
Next e_ac%
End If
End Function

Public Function combine_item0_value_and_two_line_value_and_relation(ByVal it%, _
                                   ByVal tl_n%, ByVal r_n%) As Byte
Dim tv$
Dim tn(2) As Integer
Dim con_ty As Byte
Dim rd As relation_data0_type
Dim tlv As two_line_value_data0_type
Dim temp_record As total_record_type
Dim ite As item0_data_type
If it% > 0 Then 'item0
 If item0(it%).data(0).value <> "" And item0(it%).data(0).sig = "*" Then
     If is_relation(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
          item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), _
           item0(it%).data(0).n(0), item0(it%).data(0).n(1), _
            item0(it%).data(0).n(2), item0(it%).data(0).n(3), _
             item0(it%).data(0).line_no(0), item0(it%).data(0).line_no(1), _
              tv$, tn(0), -1000, 0, 0, 0, rd, tn(1), tn(2), _
                con_ty, record_0.data0.condition_data, 0) Then
         temp_record.record_data.data0.condition_data.condition_no = 0
         Call add_record_to_record(item0(it%).data(0).record_for_value.data0.condition_data, _
                                                   temp_record.record_data.data0.condition_data)
         Call add_conditions_to_record(con_ty, tn(0), tn(1), tn(2), temp_record.record_data.data0.condition_data)
         combine_item0_value_and_two_line_value_and_relation = _
           combine_item0_value_with_relation0(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
          item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), _
           item0(it%).data(0).n(0), item0(it%).data(0).n(1), _
            item0(it%).data(0).n(2), item0(it%).data(0).n(3), _
             item0(it%).data(0).line_no(0), item0(it%).data(0).line_no(1), _
               temp_record, item0(it%).data(0).value, tv$)
            If combine_item0_value_and_two_line_value_and_relation > 1 Then
               Exit Function
            End If
     ElseIf is_two_line_value(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
          item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), item0(it%).data(0).n(0), _
           item0(it%).data(0).n(1), item0(it%).data(0).n(2), item0(it%).data(0).n(3), _
             item0(it%).data(0).line_no(0), item0(it%).data(0).line_no(1), "1", "1", "", tn(0), _
               -1000, 0, 0, 0, tlv, con_ty, record_0.data0.condition_data) Then
            If con_ty = two_line_value_ Then
             combine_item0_value_and_two_line_value_and_relation = _
              combine_item0_value_with_two_line_value0(it%, tn(0))
               If combine_item0_value_and_two_line_value_and_relation > 1 Then
                Exit Function
               End If
             End If
        ElseIf is_two_line_value(item0(it%).data(0).poi(0), item0(it%).data(0).poi(1), _
          item0(it%).data(0).poi(2), item0(it%).data(0).poi(3), item0(it%).data(0).n(0), _
           item0(it%).data(0).n(1), item0(it%).data(0).n(2), item0(it%).data(0).n(3), _
             item0(it%).data(0).line_no(0), item0(it%).data(0).line_no(1), "1", "-1", "", tn(0), _
               -1000, 0, 0, 0, tlv, con_ty, record_0.data0.condition_data) Then
            If con_ty = two_line_value_ Then
             combine_item0_value_and_two_line_value_and_relation = _
              combine_item0_value_with_two_line_value0(it%, tn(0))
               If combine_item0_value_and_two_line_value_and_relation > 1 Then
                Exit Function
               End If
             End If
     End If
 End If
ElseIf tl_n% > 0 Then
 If two_line_value(tl_n%).data(0).data0.para(0) = "1" And (two_line_value(tl_n%).data(0).data0.para(1) = "1" Or _
     two_line_value(tl_n%).data(0).data0.para(1) = "-1" Or two_line_value(tl_n%).data(0).data0.para(1) = "@1") Then
  If is_item0(two_line_value(tl_n%).data(0).data0.poi(0), two_line_value(tl_n%).data(0).data0.poi(1), _
               two_line_value(tl_n%).data(0).data0.poi(2), two_line_value(tl_n%).data(0).data0.poi(3), _
                "*", two_line_value(tl_n%).data(0).data0.n(0), two_line_value(tl_n%).data(0).data0.n(1), _
                  two_line_value(tl_n%).data(0).data0.poi(2), two_line_value(tl_n%).data(0).data0.poi(3), _
                   two_line_value(tl_n%).data(0).data0.line_no(0), two_line_value(tl_n%).data(0).data0.line_no(0), _
                    tn(0), -1000, 0, 0, "", ite) Then
           If item0(tn(0)).data(0).value <> "" Then
            combine_item0_value_and_two_line_value_and_relation = _
              combine_item0_value_with_two_line_value0(tn(0), tl_n%)
               If combine_item0_value_and_two_line_value_and_relation > 1 Then
                Exit Function
               End If
           End If
  End If
 End If
ElseIf r_n% > 0 Then
  If is_item0(Drelation(r_n%).data(0).data0.poi(0), Drelation(r_n%).data(0).data0.poi(1), _
               Drelation(r_n%).data(0).data0.poi(2), Drelation(r_n%).data(0).data0.poi(3), _
                "*", Drelation(r_n%).data(0).data0.n(0), Drelation(r_n%).data(0).data0.n(1), _
                  Drelation(r_n%).data(0).data0.n(2), Drelation(r_n%).data(0).data0.n(3), _
                   Drelation(r_n%).data(0).data0.line_no(0), Drelation(r_n%).data(0).data0.line_no(0), _
                    tn(0), -1000, 0, 0, "", ite) Then
          If item0(tn(0)).data(0).value <> "" Then
          temp_record.record_data.data0.condition_data.condition_no = 0
           Call add_conditions_to_record(relation_, r_n%, 0, 0, _
                 temp_record.record_data.data0.condition_data)
           Call add_record_to_record(item0(tn(0)).data(0).record_for_value.data0.condition_data, _
                 temp_record.record_data.data0.condition_data)
                 tv$ = Drelation(r_n%).data(0).data0.value
           combine_item0_value_and_two_line_value_and_relation = _
            set_line_value(item0(tn(0)).data(0).poi(2), item0(tn(0)).data(0).poi(3), _
                 sqr_string(divide_string(item0(tn(0)).data(0).value, tv$, False, False), _
                   True, False), item0(tn(0)).data(0).n(2), item0(tn(0)).data(0).n(3), _
                   item0(tn(0)).data(0).line_no(1), temp_record, 0, 0, False)
            If combine_item0_value_and_two_line_value_and_relation > 1 Then
               Exit Function
            End If
           combine_item0_value_and_two_line_value_and_relation = _
            set_line_value(item0(tn(0)).data(0).poi(0), item0(tn(0)).data(0).poi(1), _
                 sqr_string(time_string(item0(tn(0)).data(0).value, tv$, False, False), _
                   True, False), item0(tn(0)).data(0).n(0), item0(tn(0)).data(0).n(1), _
                   item0(tn(0)).data(0).line_no(0), temp_record, 0, 0, False)
               Exit Function
          End If
    If Drelation(r_n%).data(0).data0.poi(4) > 0 And Drelation(r_n%).data(0).data0.poi(5) > 0 Then
     If is_item0(Drelation(r_n%).data(0).data0.poi(0), Drelation(r_n%).data(0).data0.poi(1), _
               Drelation(r_n%).data(0).data0.poi(4), Drelation(r_n%).data(0).data0.poi(5), _
                "*", Drelation(r_n%).data(0).data0.n(0), Drelation(r_n%).data(0).data0.n(1), _
                  Drelation(r_n%).data(0).data0.n(4), Drelation(r_n%).data(0).data0.n(5), _
                   Drelation(r_n%).data(0).data0.line_no(0), Drelation(r_n%).data(0).data0.line_no(0), _
                    tn(0), -1000, 0, 0, "", ite) Then
          If item0(tn(0)).data(0).value <> "" Then
             tv$ = add_string("1", Drelation(r_n%).data(0).data0.value, False, False)
             tv$ = divide_string(Drelation(r_n%).data(0).data0.value, tv$, True, False)
             temp_record.record_data.data0.condition_data.condition_no = 0
             Call add_record_to_record(item0(tn(0)).data(0).record_for_value.data0.condition_data, _
                     temp_record.record_data.data0.condition_data)
             Call add_conditions_to_record(relation_, r_n%, 0, 0, temp_record.record_data.data0.condition_data)
               combine_item0_value_and_two_line_value_and_relation = _
                 combine_item0_value_with_relation0(Drelation(r_n%).data(0).data0.poi(0), _
                   Drelation(r_n%).data(0).data0.poi(1), Drelation(r_n%).data(0).data0.poi(4), _
                    Drelation(r_n%).data(0).data0.poi(5), Drelation(r_n%).data(0).data0.n(0), _
                     Drelation(r_n%).data(0).data0.n(1), Drelation(r_n%).data(0).data0.n(4), _
                      Drelation(r_n%).data(0).data0.poi(5), Drelation(r_n%).data(0).data0.line_no(0), _
                       Drelation(r_n%).data(0).data0.poi(2), temp_record, item0(tn(0)).data(0).value, _
                        tv$)
                        If combine_item0_value_and_two_line_value_and_relation > 1 Then
                           Exit Function
                        End If
          End If
     ElseIf is_item0(Drelation(r_n%).data(0).data0.poi(4), Drelation(r_n%).data(0).data0.poi(5), _
               Drelation(r_n%).data(0).data0.poi(2), Drelation(r_n%).data(0).data0.poi(3), _
                "*", Drelation(r_n%).data(0).data0.n(4), Drelation(r_n%).data(0).data0.n(5), _
                  Drelation(r_n%).data(0).data0.poi(2), Drelation(r_n%).data(0).data0.poi(3), _
                   Drelation(r_n%).data(0).data0.line_no(0), Drelation(r_n%).data(0).data0.line_no(0), _
                    tn(0), -1000, 0, 0, "", ite) Then
          If item0(tn(0)).data(0).value <> "" Then
             tv$ = add_string("1", Drelation(r_n%).data(0).data0.value, False, False)
             temp_record.record_data.data0.condition_data.condition_no = 0
             Call add_record_to_record(item0(tn(0)).data(0).record_for_value.data0.condition_data, _
                     temp_record.record_data.data0.condition_data)
             Call add_conditions_to_record(relation_, r_n%, 0, 0, temp_record.record_data.data0.condition_data)
              combine_item0_value_and_two_line_value_and_relation = _
                 combine_item0_value_with_relation0(Drelation(r_n%).data(0).data0.poi(4), _
                   Drelation(r_n%).data(0).data0.poi(5), Drelation(r_n%).data(0).data0.poi(2), _
                    Drelation(r_n%).data(0).data0.poi(3), Drelation(r_n%).data(0).data0.n(4), _
                     Drelation(r_n%).data(0).data0.n(5), Drelation(r_n%).data(0).data0.n(2), _
                      Drelation(r_n%).data(0).data0.poi(3), Drelation(r_n%).data(0).data0.line_no(2), _
                       Drelation(r_n%).data(0).data0.poi(1), temp_record, item0(tn(0)).data(0).value, _
                        tv$)
                        If combine_item0_value_and_two_line_value_and_relation > 1 Then
                           Exit Function
                        End If
          End If
     End If
    End If
  End If
 End If
End Function
Public Function combine_item0_value_with_relation0(ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, _
                       ByVal n1%, ByVal n2%, ByVal n3%, ByVal n4%, ByVal l1%, ByVal l2%, _
                         re As total_record_type, ByVal it_v$, ByVal re_v$) As Byte
Dim temp_record As total_record_type
temp_record = re
temp_record.record_data.data0.theorem_no = 1
   combine_item0_value_with_relation0 = set_line_value( _
       p1%, p2%, sqr_string(time_string(it_v$, re_v$, False, False), True, False), _
         n1%, n2%, l1%, temp_record, 0, 0, False)
         If combine_item0_value_with_relation0 > 1 Then
            Exit Function
         End If
   combine_item0_value_with_relation0 = set_line_value( _
       p3%, p4%, sqr_string(divide_string(it_v$, re_v$, False, False), True, False), _
         n3%, n4%, l2%, temp_record, 0, 0, False)
         If combine_item0_value_with_relation0 > 1 Then
            Exit Function
         End If
End Function

Public Function combine_relation_with_others_condition(no%, no_reduce As Byte) As Byte
combine_relation_with_others_condition = combine_relation_with_line_value(no%, 0, no_reduce)
If combine_relation_with_others_condition > 1 Then
 Exit Function
End If
combine_relation_with_others_condition = combine_relation_with_item(no%, no_reduce)
If combine_relation_with_others_condition > 1 Then
 Exit Function
End If
combine_relation_with_others_condition = combine_relation_with_mid_point(no%, 0, no_reduce)
If combine_relation_with_others_condition > 1 Then
 Exit Function
End If
combine_relation_with_others_condition = combine_relation_with_relation(no%, no_reduce)
 If combine_relation_with_others_condition > 1 Then
  Exit Function
 End If
combine_relation_with_others_condition = combine_relation_with_dpoint_pair(no%, 0, no_reduce)
 If combine_relation_with_others_condition > 1 Then
  Exit Function
 End If
combine_relation_with_others_condition = combine_relation_with_eline(no%, 0, no_reduce)
If combine_relation_with_others_condition > 1 Then
 Exit Function
End If
combine_relation_with_others_condition = combine_relation_with_two_line(no%, 0, no_reduce)
If combine_relation_with_others_condition > 1 Then
 Exit Function
End If
combine_relation_with_others_condition = combine_relation_with_three_line(no%, 0, no_reduce)
If combine_relation_with_others_condition > 1 Then
 Exit Function
End If
combine_relation_with_others_condition = combine_item0_value_and_two_line_value_and_relation( _
   0, 0, no%)

End Function
Public Function combine_two_v_line_value_0(ByVal v_l1%, ByVal v_l2%, ty%) As Byte
Dim tn%
Dim c_data As condition_data_type
Dim temp_record As total_record_type
Dim t_vs(1) As v_string
temp_record.record_data.data0.condition_data.condition_no = 2
temp_record.record_data.data0.condition_data.condition(1).no = v_l1%
temp_record.record_data.data0.condition_data.condition(1).ty = V_line_value_
temp_record.record_data.data0.condition_data.condition(2).no = v_l2%
temp_record.record_data.data0.condition_data.condition(2).ty = V_line_value_
If ty = 0 Then
If Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(0) = _
    Dtwo_point_line(V_line_value(v_l2%).data(0).v_line).data(0).v_poi(0) Then
     ty = 11
ElseIf Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(0) = _
    Dtwo_point_line(V_line_value(v_l2%).data(0).v_line).data(0).v_poi(1) Then
     ty = 12
ElseIf Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(1) = _
    Dtwo_point_line(V_line_value(v_l2%).data(0).v_line).data(0).v_poi(0) Then
     ty = 21
ElseIf Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(1) = _
    Dtwo_point_line(V_line_value(v_l2%).data(0).v_line).data(0).v_poi(1) Then
     ty = 22
End If
End If
 If ty = 11 Then
     combine_two_v_line_value_0 = set_V_line_value( _
        Dtwo_point_line(V_line_value(v_l2%).data(0).v_line).data(0).v_poi(1), _
         Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(1), _
          0, 0, 0, minus_string(V_line_value(v_l1%).data(0).value, _
           V_line_value(v_l2%).data(0).value, True, False), temp_record, 0, False)
     If combine_two_v_line_value_0 > 1 Then
        Exit Function
     End If
      If InStr(1, V_line_value(v_l1%).data(0).value, "x", 0) > 0 Or _
           InStr(1, V_line_value(v_l2%).data(0).value, "x", 0) > 0 Then
        If is_three_point_on_line(Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(1), _
             Dtwo_point_line(V_line_value(v_l2%).data(0).v_line).data(0).v_poi(1), _
               Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(0), _
                tn%, -1000, 0, 0, c_data, 0, 0, 0) > 0 Then
                Call add_conditions_to_record(point3_on_line_, tn%, 0, 0, _
                      temp_record.record_data.data0.condition_data)
                 If V_line_value(v_l1%).data(0).value <> "" And _
                     V_line_value(v_l2%).data(0).value <> "" Then
                    t_vs(0) = from_string_to_v_string(V_line_value(v_l1%).data(0).value)
                    t_vs(1) = from_string_to_v_string(V_line_value(v_l2%).data(0).value)
                 combine_two_v_line_value_0 = set_equation( _
                    minus_string(time_string(t_vs(0).coord(0), _
                      t_vs(1).coord(1), False, False), _
                       time_string(t_vs(1).coord(0), _
                        t_vs(0).coord(1), False, False), _
                         True, False), 0, temp_record)
                       If combine_two_v_line_value_0 > 1 Then
                        Exit Function
                       End If
                  End If
        End If
     End If
ElseIf ty = 12 Then
     combine_two_v_line_value_0 = set_V_line_value( _
        Dtwo_point_line(V_line_value(v_l2%).data(0).v_line).data(0).v_poi(0), _
         Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(1), _
          0, 0, 0, add_string(V_line_value(v_l1%).data(0).value, _
           V_line_value(v_l2%).data(0).value, True, False), temp_record, 0, False)
     If combine_two_v_line_value_0 > 1 Then
        Exit Function
     End If
     If InStr(1, V_line_value(v_l1%).data(0).value, "x", 0) > 0 Or _
           InStr(1, V_line_value(v_l2%).data(0).value, "x", 0) > 0 Then
        If is_three_point_on_line(Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(0), _
             Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(1), _
               Dtwo_point_line(V_line_value(v_l2%).data(0).v_line).data(0).v_poi(0), _
                tn%, -1000, 0, 0, c_data, 0, 0, 0) > 0 Then
                Call add_conditions_to_record(point3_on_line_, tn%, 0, 0, _
                      temp_record.record_data.data0.condition_data)
                 If V_line_value(v_l1%).data(0).value <> "" And _
                     V_line_value(v_l2%).data(0).value <> "" Then
                     t_vs(0) = from_string_to_v_string(V_line_value(v_l1%).data(0).value)
                     t_vs(1) = from_string_to_v_string(V_line_value(v_l2%).data(0).value)
                 combine_two_v_line_value_0 = set_equation( _
                    minus_string(time_string(t_vs(0).coord(0), _
                      t_vs(1).coord(1), False, False), _
                       time_string(t_vs(1).coord(0), _
                        t_vs(0).coord(1), False, False), _
                         True, False), 0, temp_record)
                       If combine_two_v_line_value_0 > 1 Then
                        Exit Function
                       End If
                  End If
        End If
     End If
 ElseIf ty = 21 Then
     combine_two_v_line_value_0 = set_V_line_value( _
        Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(0), _
         Dtwo_point_line(V_line_value(v_l2%).data(0).v_line).data(0).v_poi(1), _
          0, 0, 0, add_string(V_line_value(v_l1%).data(0).value, _
           V_line_value(v_l2%).data(0).value, True, False), temp_record, 0, False)
                        If combine_two_v_line_value_0 > 1 Then
                        Exit Function
                       End If
     If InStr(1, V_line_value(v_l1%).data(0).value, "x", 0) > 0 Or _
           InStr(1, V_line_value(v_l2%).data(0).value, "x", 0) > 0 Then
        If is_three_point_on_line(Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(1), _
             Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(0), _
               Dtwo_point_line(V_line_value(v_l2%).data(0).v_line).data(0).v_poi(1), _
                tn%, -1000, 0, 0, c_data, 0, 0, 0) > 0 Then
                Call add_conditions_to_record(point3_on_line_, tn%, 0, 0, _
                      temp_record.record_data.data0.condition_data)
                 If V_line_value(v_l1%).data(0).value <> "" And _
                     V_line_value(v_l2%).data(0).value <> "" Then
                 t_vs(0) = from_string_to_v_string(V_line_value(v_l1%).data(0).value)
                 t_vs(1) = from_string_to_v_string(V_line_value(v_l2%).data(0).value)
                 combine_two_v_line_value_0 = set_equation( _
                    minus_string(time_string(t_vs(0).coord(0), _
                      t_vs(1).coord(1), False, False), _
                       time_string(t_vs(1).coord(0), _
                        t_vs(0).coord(1), False, False), _
                         True, False), 0, temp_record)
                       If combine_two_v_line_value_0 > 1 Then
                        Exit Function
                       End If
                  End If
        End If
     End If
ElseIf ty = 22 Then
     combine_two_v_line_value_0 = set_V_line_value( _
        Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(0), _
         Dtwo_point_line(V_line_value(v_l2%).data(0).v_line).data(0).v_poi(0), _
          0, 0, 0, minus_string(V_line_value(v_l1%).data(0).value, _
           V_line_value(v_l2%).data(0).value, True, False), temp_record, 0, False)
     If combine_two_v_line_value_0 > 1 Then
        Exit Function
     End If
      If InStr(1, V_line_value(v_l1%).data(0).value, "x", 0) > 0 Or _
           InStr(1, V_line_value(v_l2%).data(0).value, "x", 0) > 0 Then
        If is_three_point_on_line(Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(1), _
             Dtwo_point_line(V_line_value(v_l2%).data(0).v_line).data(0).v_poi(0), _
               Dtwo_point_line(V_line_value(v_l1%).data(0).v_line).data(0).v_poi(0), _
                tn%, -1000, 0, 0, c_data, 0, 0, 0) > 0 Then
                Call add_conditions_to_record(point3_on_line_, tn%, 0, 0, _
                      temp_record.record_data.data0.condition_data)
                 If V_line_value(v_l1%).data(0).value <> "" And _
                     V_line_value(v_l2%).data(0).value <> "0" Then
                 t_vs(0) = from_string_to_v_string(V_line_value(v_l1%).data(0).value)
                 t_vs(1) = from_string_to_v_string(V_line_value(v_l2%).data(0).value)
                 combine_two_v_line_value_0 = set_equation( _
                    minus_string(time_string(t_vs(0).coord(0), _
                      t_vs(1).coord(1), False, False), _
                       time_string(t_vs(1).coord(0), _
                        t_vs(0).coord(1), False, False), _
                         True, False), 0, temp_record)
                       If combine_two_v_line_value_0 > 1 Then
                        Exit Function
                       End If
                  End If
        End If
     End If
End If
End Function

Function combine_v_line_value_with_item0(no%, ByVal i%)
Dim tn%
Dim val As String
Dim v_lv As V_line_value_data0_type
Dim t_v_val As String
Dim c_data As condition_data_type
Dim temp_record As total_record_type
 temp_record.record_data.data0.condition_data.condition_no = 1
 temp_record.record_data.data0.condition_data.condition(1).ty = V_line_value_
 temp_record.record_data.data0.condition_data.condition(1).no = no%
   If item0(i%).data(0).poi(1) = -10 Then
      If item0(i%).data(0).poi(0) = V_line_value(no%).data(0).v_line Then
       If item0(i%).data(0).poi(2) > 0 Then
          If is_V_line_value(Dtwo_point_line(item0(i%).data(0).poi(2)).data(0).v_poi(0), _
              Dtwo_point_line(item0(i%).data(0).poi(2)).data(0).v_poi(1), 0, 0, 0, t_v_val, _
               tn%, -1000, 0, 0, 0, v_lv, False) Then
                 Call add_conditions_to_record(V_line_value_, tn%, 0, 0, _
                               temp_record.record_data.data0.condition_data)
                val = time_string(V_line_value(no%).data(0).value, _
                  V_line_value(tn%).data(0).value, True, False)
                    combine_v_line_value_with_item0 = _
                     set_item0_value(i%, 0, 0, "", "", "", val, 0, _
                            temp_record.record_data.data0.condition_data)
                    If combine_v_line_value_with_item0 > 1 Then
                       Exit Function
                    End If
          Else
                    combine_v_line_value_with_item0 = set_item0(item0(i%).data(0).poi(2), _
                       item0(i%).data(0).poi(3), 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", _
                         V_line_value(no%).data(0).value, "", "1", 0, c_data, _
                         i%, 0, 0, 0, temp_record.record_data.data0.condition_data, False)
                    If combine_v_line_value_with_item0 > 1 Then
                       Exit Function
                    End If
          End If
       Else
                    combine_v_line_value_with_item0 = _
                     set_item0_value(i%, 0, 0, "", "", "", _
                        V_line_value(no%).data(0).value, 0, temp_record.record_data.data0.condition_data)
                    If combine_v_line_value_with_item0 > 1 Then
                       Exit Function
                    End If
       End If
     End If
   ElseIf item0(i%).data(0).poi(3) = -10 Then
      If item0(i%).data(0).poi(2) = V_line_value(no%).data(0).v_line Then
       If item0(i%).data(0).poi(0) > 0 Then
          If is_V_line_value(Dtwo_point_line(item0(i%).data(0).poi(0)).data(0).v_poi(0), _
              Dtwo_point_line(item0(i%).data(0).poi(0)).data(0).v_poi(1), 0, 0, 0, t_v_val, _
               tn%, -1000, 0, 0, 0, v_lv, False) Then
                 Call add_conditions_to_record(V_line_value_, tn%, 0, 0, _
                                temp_record.record_data.data0.condition_data)
                val = time_string(V_line_value(no%).data(0).value, _
                  V_line_value(tn%).data(0).value, True, False)
                    combine_v_line_value_with_item0 = _
                     set_item0_value(i%, 0, 0, "", "", "", val, 0, _
                            temp_record.record_data.data0.condition_data)
                    If combine_v_line_value_with_item0 > 1 Then
                       Exit Function
                    End If
          Else
                    combine_v_line_value_with_item0 = set_item0(item0(i%).data(0).poi(0), _
                       item0(i%).data(0).poi(1), 0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", _
                         V_line_value(no%).data(0).value, "", "1", 0, c_data, i%, 0, 0, 0, _
                          temp_record.record_data.data0.condition_data, False)
                    If combine_v_line_value_with_item0 > 1 Then
                       Exit Function
                    End If
          End If
      Else
                    combine_v_line_value_with_item0 = _
                     set_item0_value(i%, 0, 0, "", "", "", _
                        V_line_value(no%).data(0).value, 0, _
                            temp_record.record_data.data0.condition_data)
                    If combine_v_line_value_with_item0 > 1 Then
                       Exit Function
                    End If
      End If
     End If
   End If
End Function
Function combine_v_line_value_with_item(no%) As Byte
Dim i%
Dim val As String
Dim v_val As v_string
Dim temp_record As total_record_type
For i% = 1 To last_conditions.last_cond(1).item0_no
  combine_v_line_value_with_item = combine_v_line_value_with_item0(no%, i%)
     If combine_v_line_value_with_item > 1 Then
        Exit Function
     End If
Next i%
End Function
Public Function combine_v_line_value_with_v_relation(v_lv%) As Byte
Dim i%
For i% = 1 To last_conditions.last_cond(1).v_relation_no
  If v_Drelation(i%).data(0).data0.v_line(0) = V_line_value(v_lv%).data(0).v_line Then
    combine_v_line_value_with_v_relation = _
      combine_v_line_value_with_v_relation0(v_lv%, i%, 0)
     If combine_v_line_value_with_v_relation > 1 Then
        Exit Function
     End If
  ElseIf v_Drelation(i%).data(0).data0.v_line(1) = V_line_value(v_lv%).data(0).v_line Then
    combine_v_line_value_with_v_relation = _
      combine_v_line_value_with_v_relation0(v_lv%, i%, 1)
     If combine_v_line_value_with_v_relation > 1 Then
        Exit Function
     End If
  ElseIf v_Drelation(i%).data(0).data0.v_line(2) = V_line_value(v_lv%).data(0).v_line Then
    combine_v_line_value_with_v_relation = _
      combine_v_line_value_with_v_relation0(v_lv%, i%, 2)
     If combine_v_line_value_with_v_relation > 1 Then
        Exit Function
     End If
  End If
Next i%
End Function
Public Function combine_v_relation_with_v_line_value(v_re%) As Byte
Dim i%
For i% = 1 To last_conditions.last_cond(1).v_line_value_no
  If v_Drelation(v_re%).data(0).data0.v_line(0) = V_line_value(i%).data(0).v_line Then
    combine_v_relation_with_v_line_value = _
      combine_v_line_value_with_v_relation0(i%, v_re%, 0)
     If combine_v_relation_with_v_line_value > 1 Then
        Exit Function
     End If
  ElseIf v_Drelation(v_re%).data(0).data0.v_line(1) = V_line_value(i%).data(0).v_line Then
    combine_v_relation_with_v_line_value = _
      combine_v_line_value_with_v_relation0(i%, v_re%, 1)
     If combine_v_relation_with_v_line_value > 1 Then
        Exit Function
     End If
  ElseIf v_Drelation(v_re%).data(0).data0.v_line(2) = V_line_value(i%).data(0).v_line Then
    combine_v_relation_with_v_line_value = _
      combine_v_line_value_with_v_relation0(i%, v_re%, 2)
     If combine_v_relation_with_v_line_value > 1 Then
        Exit Function
     End If
  
  End If
Next i%
End Function
Public Function combine_v_line_value_with_v_relation0(v_lv%, re_no%, ty As Byte) As Byte
Dim temp_record As total_record_type
Dim tv$
temp_record.record_data.data0.condition_data.condition_no = 2
temp_record.record_data.data0.condition_data.condition(1).ty = V_line_value_
temp_record.record_data.data0.condition_data.condition(1).no = v_lv%
temp_record.record_data.data0.condition_data.condition(2).ty = v_relation_
temp_record.record_data.data0.condition_data.condition(2).no = re_no%
If ty = 0 Then
   combine_v_line_value_with_v_relation0 = set_V_line_value( _
      Dtwo_point_line(v_Drelation(re_no%).data(0).data0.v_line(1)).data(0).v_poi(0), _
       Dtwo_point_line(v_Drelation(re_no%).data(0).data0.v_line(1)).data(0).v_poi(1), _
           0, 0, 0, divide_string(V_line_value(v_lv%).data(0).value, _
            v_Drelation(re_no%).data(0).data0.value, True, False), temp_record, 0, False)
     If combine_v_line_value_with_v_relation0 > 1 Then
        Exit Function
     End If
     If v_Drelation(re_no%).data(0).data0.v_line(2) > 0 Then
     tv$ = divide_string(add_string("1", v_Drelation(re_no%).data(0).data0.value, False, False), _
            v_Drelation(re_no%).data(0).data0.value, True, False)
      combine_v_line_value_with_v_relation0 = set_V_line_value( _
        Dtwo_point_line(v_Drelation(re_no%).data(0).data0.v_line(2)).data(0).v_poi(0), _
          Dtwo_point_line(v_Drelation(re_no%).data(0).data0.v_line(2)).data(0).v_poi(1), _
            0, 0, 0, time_string(V_line_value(v_lv%).data(0).value, _
              tv$, True, False), temp_record, 0, False)
     If combine_v_line_value_with_v_relation0 > 1 Then
        Exit Function
     End If
     End If
ElseIf ty = 1 Then
   combine_v_line_value_with_v_relation0 = set_V_line_value( _
      Dtwo_point_line(v_Drelation(re_no%).data(0).data0.v_line(0)).data(0).v_poi(0), _
       Dtwo_point_line(v_Drelation(re_no%).data(0).data0.v_line(0)).data(0).v_poi(1), _
          0, 0, 0, time_string(V_line_value(v_lv%).data(0).value, _
            v_Drelation(re_no%).data(0).data0.value, True, False), temp_record, 0, False)
     If combine_v_line_value_with_v_relation0 > 1 Then
        Exit Function
     End If
     If v_Drelation(re_no%).data(0).data0.v_line(2) > 0 Then
     tv$ = add_string("1", v_Drelation(re_no%).data(0).data0.value, True, False)
      combine_v_line_value_with_v_relation0 = set_V_line_value( _
        Dtwo_point_line(v_Drelation(re_no%).data(0).data0.v_line(2)).data(0).v_poi(0), _
          Dtwo_point_line(v_Drelation(re_no%).data(0).data0.v_line(2)).data(0).v_poi(1), _
            0, 0, 0, time_string(V_line_value(v_lv%).data(0).value, _
              tv$, True, False), temp_record, 0, False)
     If combine_v_line_value_with_v_relation0 > 1 Then
        Exit Function
     End If
     
     End If
ElseIf ty = 2 Then
   tv$ = add_string("1", v_Drelation(re_no%).data(0).data0.value, True, False)
   combine_v_line_value_with_v_relation0 = set_V_line_value( _
      Dtwo_point_line(v_Drelation(re_no%).data(0).data0.v_line(1)).data(0).v_poi(0), _
       Dtwo_point_line(v_Drelation(re_no%).data(0).data0.v_line(1)).data(0).v_poi(1), _
          0, 0, 0, divide_string(V_line_value(v_lv%).data(0).value, _
            tv$, True, False), temp_record, 0, False)
     If combine_v_line_value_with_v_relation0 > 1 Then
        Exit Function
     End If
   tv$ = divide_string(v_Drelation(re_no%).data(0).data0.value, tv$, True, False)
   combine_v_line_value_with_v_relation0 = set_V_line_value( _
      Dtwo_point_line(v_Drelation(re_no%).data(0).data0.v_line(0)).data(0).v_poi(0), _
       Dtwo_point_line(v_Drelation(re_no%).data(0).data0.v_line(0)).data(0).v_poi(1), _
          0, 0, 0, time_string(V_line_value(v_lv%).data(0).value, _
            tv$, True, False), temp_record, 0, False)
     If combine_v_line_value_with_v_relation0 > 1 Then
        Exit Function
     End If
End If
End Function
Public Function combine_v_relation_with_v_relation(re_no%) As Byte
Dim i%
Dim tv$
For i% = 1 To re_no% - 1
 If v_Drelation(re_no%).data(0).data0.v_line(0) = _
       v_Drelation(i%).data(0).data0.v_line(0) Then
     combine_v_relation_with_v_relation = _
         combine_v_relation_with_v_relation0(re_no%, i%, 0)
               If combine_v_relation_with_v_relation > 1 Then
                  Exit Function
               End If
 ElseIf v_Drelation(re_no%).data(0).data0.v_line(0) = _
       v_Drelation(i%).data(0).data0.v_line(1) Then
     combine_v_relation_with_v_relation = _
         combine_v_relation_with_v_relation0(re_no%, i%, 1)
               If combine_v_relation_with_v_relation > 1 Then
                  Exit Function
               End If
 ElseIf v_Drelation(re_no%).data(0).data0.v_line(0) = _
       v_Drelation(i%).data(0).data0.v_line(2) Then
     combine_v_relation_with_v_relation = _
         combine_v_relation_with_v_relation0(re_no%, i%, 2)
               If combine_v_relation_with_v_relation > 1 Then
                  Exit Function
               End If
 ElseIf v_Drelation(re_no%).data(0).data0.v_line(1) = _
       v_Drelation(i%).data(0).data0.v_line(0) Then
        combine_v_relation_with_v_relation = _
         combine_v_relation_with_v_relation0(i%, re_no%, 1)
               If combine_v_relation_with_v_relation > 1 Then
                  Exit Function
               End If
ElseIf v_Drelation(re_no%).data(0).data0.v_line(1) = _
       v_Drelation(i%).data(0).data0.v_line(1) Then
        combine_v_relation_with_v_relation = _
         combine_v_relation_with_v_relation0(re_no%, i%, 11)
               If combine_v_relation_with_v_relation > 1 Then
                  Exit Function
               End If
  ElseIf v_Drelation(re_no%).data(0).data0.v_line(1) = _
       v_Drelation(i%).data(0).data0.v_line(2) Then
        combine_v_relation_with_v_relation = _
         combine_v_relation_with_v_relation0(re_no%, i%, 12)
               If combine_v_relation_with_v_relation > 1 Then
                  Exit Function
               End If
ElseIf v_Drelation(re_no%).data(0).data0.v_line(2) = _
       v_Drelation(i%).data(0).data0.v_line(0) Then
        combine_v_relation_with_v_relation = _
         combine_v_relation_with_v_relation0(i%, re_no%, 2)
               If combine_v_relation_with_v_relation > 1 Then
                  Exit Function
               End If
ElseIf v_Drelation(re_no%).data(0).data0.v_line(2) = _
       v_Drelation(i%).data(0).data0.v_line(1) Then
        combine_v_relation_with_v_relation = _
         combine_v_relation_with_v_relation0(i%, re_no%, 12)
               If combine_v_relation_with_v_relation > 1 Then
                  Exit Function
               End If
ElseIf v_Drelation(re_no%).data(0).data0.v_line(2) = _
       v_Drelation(i%).data(0).data0.v_line(2) And _
         v_Drelation(i%).data(0).data0.v_line(2) > 0 Then
        combine_v_relation_with_v_relation = _
         combine_v_relation_with_v_relation0(i%, re_no%, 22)
               If combine_v_relation_with_v_relation > 1 Then
                  Exit Function
               End If
End If
Next i%
End Function
Public Function combine_v_relation_with_v_relation0(re_no1%, re_no2%, ty As Byte) As Byte
Dim temp_record As total_record_type
Dim tv$
temp_record.record_data.data0.condition_data.condition_no = 2
temp_record.record_data.data0.condition_data.condition(1).ty = v_relation_
temp_record.record_data.data0.condition_data.condition(1).no = re_no1%
temp_record.record_data.data0.condition_data.condition(2).ty = v_relation_
temp_record.record_data.data0.condition_data.condition(2).no = re_no2%
If ty = 0 Then '00
     combine_v_relation_with_v_relation0 = set_v_relation( _
      Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(1)).data(0).v_poi(0), _
       Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(1)).data(0).v_poi(1), _
        Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(1)).data(0).v_poi(0), _
         Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(1)).data(0).v_poi(1), _
          divide_string(v_Drelation(re_no2%).data(0).data0.value, _
             v_Drelation(re_no1%).data(0).data0.value, True, False), 0, _
              temp_record)
               If combine_v_relation_with_v_relation0 > 1 Then
                  Exit Function
               End If
ElseIf ty = 1 Then '01
    combine_v_relation_with_v_relation0 = set_v_relation( _
      Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(0)).data(0).v_poi(0), _
       Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(0)).data(0).v_poi(1), _
        Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(1)).data(0).v_poi(0), _
         Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(1)).data(0).v_poi(1), _
          time_string(v_Drelation(re_no1%).data(0).data0.value, _
             v_Drelation(re_no2%).data(0).data0.value, True, False), 0, _
              temp_record)
               If combine_v_relation_with_v_relation0 > 1 Then
                  Exit Function
               End If
ElseIf ty = 11 Then '11
      combine_v_relation_with_v_relation0 = set_v_relation( _
      Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(0)).data(0).v_poi(0), _
       Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(0)).data(0).v_poi(1), _
        Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(0)).data(0).v_poi(0), _
         Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(0)).data(0).v_poi(1), _
          divide_string(v_Drelation(re_no1%).data(0).data0.value, _
             v_Drelation(re_no2%).data(0).data0.value, True, False), 0, _
              temp_record)
               If combine_v_relation_with_v_relation0 > 1 Then
                  Exit Function
               End If
ElseIf ty = 2 Then '02
    tv$ = time_string(v_Drelation(re_no1%).data(0).data0.value, _
             v_Drelation(re_no2%).data(0).data0.value, True, False)
    tv$ = divide_string(tv$, add_string("1", v_Drelation(re_no2%).data(0).data0.value, False, False), True, False)
    combine_v_relation_with_v_relation0 = set_v_relation( _
      Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(0)).data(0).v_poi(0), _
       Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(0)).data(0).v_poi(1), _
        Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(1)).data(0).v_poi(0), _
         Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(1)).data(0).v_poi(1), _
          tv$, 0, temp_record)
               If combine_v_relation_with_v_relation0 > 1 Then
                  Exit Function
               End If
    tv$ = divide_string(v_Drelation(re_no1%).data(0).data0.value, add_string("1", v_Drelation(re_no2%).data(0).data0.value, False, False), True, False)
    combine_v_relation_with_v_relation0 = set_v_relation( _
      Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(1)).data(0).v_poi(0), _
       Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(1)).data(0).v_poi(1), _
        Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(1)).data(0).v_poi(0), _
         Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(1)).data(0).v_poi(1), _
          tv$, 0, temp_record)
               If combine_v_relation_with_v_relation0 > 1 Then
                  Exit Function
               End If
ElseIf ty = 12 Then '12
      combine_v_relation_with_v_relation0 = set_v_relation( _
      Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(0)).data(0).v_poi(0), _
       Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(0)).data(0).v_poi(1), _
        Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(1)).data(0).v_poi(0), _
         Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(1)).data(0).v_poi(1), _
          time_string(v_Drelation(re_no1%).data(0).data0.value, _
             add_string(v_Drelation(re_no2%).data(0).data0.value, "1", False, False), _
              True, False), 0, temp_record)
               If combine_v_relation_with_v_relation0 > 1 Then
                  Exit Function
               End If
    tv$ = add_string(divide_string("1", v_Drelation(re_no2%).data(0).data0.value, False, False), _
                "1", True, False)
    tv$ = time_string(tv$, v_Drelation(re_no1%).data(0).data0.value, _
                        True, False)
    combine_v_relation_with_v_relation0 = set_v_relation( _
      Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(0)).data(0).v_poi(0), _
       Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(0)).data(0).v_poi(1), _
        Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(0)).data(0).v_poi(0), _
         Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(0)).data(0).v_poi(1), _
          tv$, 0, temp_record)
               If combine_v_relation_with_v_relation0 > 1 Then
                  Exit Function
               End If
ElseIf ty = 22 Then '22
    tv$ = divide_string(v_Drelation(re_no1%).data(0).data0.value, _
        add_string("1", v_Drelation(re_no1%).data(0).data0.value, False, False), True, False)
      combine_v_relation_with_v_relation0 = set_v_relation( _
      Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(0)).data(0).v_poi(0), _
       Dtwo_point_line(v_Drelation(re_no1%).data(0).data0.v_line(0)).data(0).v_poi(1), _
        Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(1)).data(0).v_poi(0), _
         Dtwo_point_line(v_Drelation(re_no2%).data(0).data0.v_line(1)).data(0).v_poi(1), _
          time_string(tv$, _
             add_string(v_Drelation(re_no2%).data(0).data0.value, "1", False, False), _
              True, False), 0, temp_record)
               If combine_v_relation_with_v_relation0 > 1 Then
                  Exit Function
               End If
End If
End Function
Public Function combine_item0_with_two_line_value_(ByVal it%, ByVal tl%, ByVal i%, ByVal j%) As Byte
Dim i_%, j_%
Dim temp_record As total_record_type
Dim it_(1) As Integer
Dim para(1) As String
Dim c_data0 As condition_data_type
Dim tv$
Dim re As record_type0
Call add_conditions_to_record(two_line_value_, tl%, 0, 0, temp_record.record_data.data0.condition_data)
'Call add_conditions_to_record(item0_, it%, 0, 0, temp_record.record_data.data0.condition_data)
i_% = (i% + 1) Mod 2
j_% = (j% + 1) Mod 2
If item0(it%).data(0).sig = "*" Then
     Call set_item0(item0(it%).data(0).poi(2 * i_%), item0(it%).data(0).poi(2 * i_% + 1), _
            two_line_value(tl%).data(0).data0.poi(2 * j_%), _
              two_line_value(tl%).data(0).data0.poi(2 * j_% + 1), _
                 "*", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), 0, _
                    c_data0, 0, it_(0), 0, 0, temp_record.record_data.data0.condition_data, False)
      Call set_item0(item0(it%).data(0).poi(2 * i_%), item0(it%).data(0).poi(2 * i_% + 1), _
            0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), 0, _
                    c_data0, 0, it_(1), 0, 0, temp_record.record_data.data0.condition_data, False)
      If item0(it%).data(0).value = "" Then
       If it_(0) = 0 And it_(1) = 0 Then
         tv$ = time_string(two_line_value(tl%).data(0).data0.value, para(1), False, False)
         tv$ = minus_string(tv$, _
            time_string(two_line_value(tl%).data(0).data0.para(j_%), para(0), False, False), _
                False, False)
         tv$ = divide_string(tv$, two_line_value(tl%).data(0).data0.para(j%), True, False)
         combine_item0_with_two_line_value_ = set_item0_value(it%, 0, 0, "", "", "", tv$, 0, _
            temp_record.record_data.data0.condition_data)
            If combine_item0_with_two_line_value_ > 1 Then
               Exit Function
            End If
       ElseIf (it_(0) > 0 And it_(1) = 0) Or (it_(0) = 0 And it_(1) > 0) Then
          combine_item0_with_two_line_value_ = add_new_item_to_item(it_(0), it_(1), _
            divide_string(time_string("-1", time_string(two_line_value(tl%).data(0).data0.para(j_%), _
              para(0), True, False), False, False), two_line_value(tl%).data(0).data0.para(j%), True, False), _
                divide_string(time_string(two_line_value(tl%).data(0).data0.value, _
                  para(1), False, False), two_line_value(tl%).data(0).data0.para(j%), True, False), _
                    it%, temp_record.record_data.data0.condition_data)
              If combine_item0_with_two_line_value_ > 1 Then
               Exit Function
              End If
       End If
      Else
       If (it_(0) = 0 And it_(1) > 0) Or (it_(0) > 0 And it_(1) > 0) Then
         combine_item0_with_two_line_value_ = set_general_string(it_(0), it_(1), 0, 0, _
             time_string(time_string(para(0), "-1", False, False), two_line_value(tl%).data(0).data0.para(j_%), _
               True, False), time_string(para(1), two_line_value(tl%).data(0).data0.value, True, False), _
                "0", "0", time_string(item0(it).data(0).value, two_line_value(tl%).data(0).data0.para(j%), _
                   True, False), 0, 0, 0, temp_record, 0, 0)
            If combine_item0_with_two_line_value_ > 1 Then
               Exit Function
            End If
       End If
      End If
   ElseIf item0(it%).data(0).sig = "~" Then
       If item0(it%).data(0).value = "" Then
        Call set_item0(two_line_value(tl%).data(0).data0.poi(2 * j_%), _
                       two_line_value(tl%).data(0).data0.poi(2 * j_% + 1), _
                        0, 0, "~", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), 0, _
                         c_data0, 0, it_(0), 0, 0, _
                            temp_record.record_data.data0.condition_data, False)
        it_(1) = 0
         If it_(0) > 0 Then
          combine_item0_with_two_line_value_ = add_new_item_to_item(it_(0), 0, _
            divide_string(time_string(time_string("-1", para(0), False, False), _
              two_line_value(tl%).data(0).data0.para(j_%), False, False), _
               two_line_value(tl%).data(0).data0.para(j%), True, False), _
              divide_string(two_line_value(tl%).data(0).data0.value, _
               two_line_value(tl%).data(0).data0.para(j%), True, False), _
                   it%, temp_record.record_data.data0.condition_data)
              If combine_item0_with_two_line_value_ > 1 Then
               Exit Function
              End If
          Else
           combine_item0_with_two_line_value_ = set_item0_value(it%, 0, 0, "", "", "", _
              divide_string(minus_string(two_line_value(tl%).data(0).data0.value, _
                  time_string(two_line_value(tl%).data(0).data0.para(j_%), para(0), False, False), _
                   False, False), two_line_value(tl%).data(0).data0.para(j%), True, False), _
                    0, temp_record.record_data.data0.condition_data)
              If combine_item0_with_two_line_value_ > 1 Then
               Exit Function
              End If
          End If
        Else
         combine_item0_with_two_line_value_ = set_line_value(two_line_value(tl%).data(0).data0.poi(2 * j_%), _
             two_line_value(tl%).data(0).data0.poi(2 * j_% + 1), divide_string(minus_string( _
                two_line_value(tl%).data(0).data0.value, time_string(item0(it%).data(0).value, _
                  two_line_value(tl%).data(0).data0.para(j%), False, False), False, False), _
                   two_line_value(tl%).data(0).data0.para(j_%), True, False), 0, 0, 0, _
                     temp_record, 0, 0, False)
              If combine_item0_with_two_line_value_ > 1 Then
               Exit Function
              End If
        End If
ElseIf item0(it%).data(0).sig = "/" Then
   If item0(it%).data(0).value <> "" Then
      If i% = 0 Then
       combine_item0_with_two_line_value_ = set_two_line_value(two_line_value(tl%).data(0).data0.poi(2 * j_%), _
              two_line_value(tl%).data(0).data0.poi(2 * j_% + 1), item0(it%).data(0).poi(2 * i_%), _
                item0(it%).data(0).poi(2 * i_% + 1), 0, 0, 0, 0, 0, 0, two_line_value(tl%).data(0).data0.para(j_%), _
                 divide_string(two_line_value(tl%).data(0).data0.para(j%), item0(it%).data(0).value, _
                  True, False), two_line_value(tl%).data(0).data0.value, temp_record, 0, 0)
              If combine_item0_with_two_line_value_ > 1 Then
               Exit Function
              End If
      ElseIf i% = 1 Then
       combine_item0_with_two_line_value_ = set_two_line_value(two_line_value(tl%).data(0).data0.poi(2 * j_%), _
              two_line_value(tl%).data(0).data0.poi(2 * j_% + 1), item0(it%).data(0).poi(2 * i_%), _
                item0(it%).data(0).poi(2 * i_% + 1), 0, 0, 0, 0, 0, 0, two_line_value(tl%).data(0).data0.para(j_%), _
                 time_string(two_line_value(tl%).data(0).data0.para(j%), item0(it%).data(0).value, _
                  True, False), two_line_value(tl%).data(0).data0.value, temp_record, 0, 0)
              If combine_item0_with_two_line_value_ > 1 Then
               Exit Function
              End If
      End If
   ElseIf i% = 0 Then
     Call set_item0(item0(it%).data(0).poi(2 * i_%), item0(it%).data(0).poi(2 * i_% + 1), _
            two_line_value(tl%).data(0).data0.poi(2 * j_%), _
              two_line_value(tl%).data(0).data0.poi(2 * j_% + 1), _
                 "/", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(0), 0, _
                    c_data0, 0, it_(0), 0, 0, temp_record.record_data.data0.condition_data, False)
     Call set_item0(0, 0, item0(it%).data(0).poi(2 * i_%), item0(it%).data(0).poi(2 * i_% + 1), _
             "/", 0, 0, 0, 0, 0, 0, "1", "1", "1", "", para(1), 0, _
                    c_data0, 0, it_(1), 0, 0, temp_record.record_data.data0.condition_data, False)
     If it_(0) = 0 And it_(1) = 0 Then
        combine_item0_with_two_line_value_ = set_item0_value(it%, 0, 0, "", "", "", _
                   divide_string(minus_string( _
                    time_string(para(1), two_line_value(tl%).data(0).data0.value, False, False), _
                       time_string(para(0), two_line_value(tl%).data(0).data0.para(j_%), False, False), _
                        False, False), two_line_value(tl%).data(0).data0.para(j%), True, False), _
                          0, temp_record.record_data.data0.condition_data)
               If combine_item0_with_two_line_value_ > 1 Then
                Exit Function
               End If
     ElseIf (it_(0) = 0 And it_(1) > 0) Or (it_(0) > 0 And it_(1) = 0) Then
          combine_item0_with_two_line_value_ = add_new_item_to_item(it_(0), it_(1), _
            divide_string(time_string("-1", time_string(two_line_value(tl%).data(0).data0.para(j_%), _
              para(0), True, False), False, False), two_line_value(tl%).data(0).data0.para(j%), True, False), _
                divide_string(time_string(two_line_value(tl%).data(0).data0.value, _
                  para(1), False, False), two_line_value(tl%).data(0).data0.para(j%), True, False), _
                    it%, temp_record.record_data.data0.condition_data)
              If combine_item0_with_two_line_value_ > 1 Then
               Exit Function
              End If
     End If
   End If
End If
End Function
Public Function combine_para_with_element(ByVal para$, ByVal ele As String, _
                                  initial_string As String, dis_ty As Byte) As String
'dis=0 dis=1 parint
Dim pA$
Dim sig As String
If para = "1" Then
 sig = "+"
 para = ""
ElseIf para = "-1" Then
 sig = "-"
 para = ""
ElseIf para = "0" Or para = "-0" Or para = "+0" Then
 If ele = "1" Then
  sig = "+"
  para = "0"
  ele = ""
 Else
  sig = ""
  para = ""
 End If
Else
 If Mid$(para$, 1, 1) = "-" Then
    para$ = time_string("-1", para$, True, False)
     sig = "-"
 Else
    sig = "+"
 End If
End If
 If ele = "1" Then
    ele = ""
 End If
 If (InStr(2, para$, "+", 0) > 0 Or InStr(2, para$, "-", 0) > 0) And InStr(2, para$, "/", 0) = 0 And ele <> "" Then
       pA$ = "(" + display_string_(para$, dis_ty) + ")"
 ElseIf para$ <> "" Then
       pA$ = display_string_(para$, dis_ty)
 End If
If initial_string = "" Then
   combine_para_with_element = pA$ + ele
   If sig = "-" Then
    combine_para_with_element = "-" + combine_para_with_element
   End If
Else
   If para$ <> "0" Then
    combine_para_with_element = initial_string + sig + pA$ + ele
   Else
    combine_para_with_element = initial_string
   End If
End If
End Function

Attribute VB_Name = "aid_condition_moudle"
Option Explicit
Public Function set_add_aid_point_for_two_line(p%, l1%, paral_or_verti_ As Integer, _
                     l2%, n%) As Boolean 'paral_or_verti =1 =paral 0=verti_,p%=0,两线交点
                   '添加过p%点平行垂直l1的直线与直线l2的交点
Dim i%, tn%, tp%
Dim t_l(1) As add_point_for_two_line_type
'On Error GoTo set_add_aid_point_for_two_line_error
If is_add_point_for_two_line(p%, paral_or_verti_, _
                                             l1%, 0, 0, l2%) Then '已添加此类辅助点
     Exit Function
Else
    n% = set_add_aid_point_for_two_line_(p%, paral_or_verti_, l1, 0, 0, l2%)
       set_add_aid_point_for_two_line = True
End If
End Function

Private Function is_add_point_for_two_line(ByVal p1%, ByVal ty1 As Integer, _
                              ByVal l1%, ByVal p2%, ByVal ty2 As Integer, ByVal l2%) As Boolean
Dim i%
Call simple_line_from_point_line(p1%, ty1, l1%)
Call simple_line_from_point_line(p2%, ty2, l2%)
'*******************************************************************
If p1% = 0 And p2% = 0 Then
   If l1% > l2% Then
   Call exchange_two_integer(l1%, l2%)
   Call exchange_two_integer(p1%, p2%)
   Call exchange_two_integer(ty1%, ty2)
   End If
ElseIf p1% > p2% Then
   Call exchange_two_integer(l1%, l2%)
   Call exchange_two_integer(p1%, p2%)
   Call exchange_two_integer(ty1%, ty2)
End If
If p1% = 0 And p2% = 0 Then
   If is_line_line_intersect(l1%, l2%, 0, 0, 0) > 0 Then
       is_add_point_for_two_line = True
        Exit Function
   End If
End If
'*********************************************************************************
For i% = 1 To last_add_aid_point_for_two_line
    If is_same_line_from_point_line(p1%, ty1, l1%, _
                add_aid_point_for_two_line_(i%).s_poi(0), _
                 add_aid_point_for_two_line_(i%).paral_or_verti(0), _
                  add_aid_point_for_two_line_(i%).line_no(0)) Then
       If is_same_line_from_point_line(p2%, ty2, l2%, _
                  add_aid_point_for_two_line_(i%).s_poi(1), _
                   add_aid_point_for_two_line_(i%).paral_or_verti(1), _
                    add_aid_point_for_two_line_(i%).line_no(1)) Then
              is_add_point_for_two_line = True
               Exit Function
       End If
    End If
Next i%
End Function
Private Function is_same_line_from_point_line(ByVal p1%, ByVal ty1%, ByVal l1%, _
                 ByVal p2%, ByVal ty2%, ByVal l2%) As Boolean
Dim tl%
    If p1% = 0 And p2% = 0 Then
       If l1% = l2% Then
          is_same_line_from_point_line = True
       End If
    ElseIf p1% = p2% Then
    If is_dparal(l2%, l1%, 0, -1000, 0, 0, 0, 0) Then
         tl% = line_number0(p1%, p2%, 0, 0)
       If ty1 = ty2 Then
          is_same_line_from_point_line = True
       End If
    ElseIf is_dverti(l2%, l1%, 0, -1000, 0, 0, 0, 0) Then
       If (ty1 = paral_ And ty2 = verti_) Or (ty2 = paral_ And ty1 = verti_) Then
           is_same_line_from_point_line = True
       End If
    End If
    End If
End Function

Private Function set_add_aid_point_for_two_line_(ByVal p1%, ByVal ty1%, ByVal l1%, _
                                        ByVal p2%, ByVal ty2%, ByVal l2%) As Integer
 Call simple_line_from_point_line(p1%, ty1%, l1%)
 Call simple_line_from_point_line(p2%, ty2%, l2%)
If p1% = 0 And p2% = 0 Then
   If l1% > l2% Then
   Call exchange_two_integer(l1%, l2%)
   Call exchange_two_integer(p1%, p2%)
   Call exchange_two_integer(ty1%, ty2)
   End If
ElseIf p1% > p2% Then
   Call exchange_two_integer(l1%, l2%)
   Call exchange_two_integer(p1%, p2%)
   Call exchange_two_integer(ty1%, ty2)
End If
If ty1 = paral_ Then
 If is_point_in_line3(p1%, m_lin(l1%).data(0).data0, 0) Then
    p1% = 0
    ty1 = 0
 End If
End If
If ty2 = paral_ Then
 If is_point_in_line3(p2%, m_lin(l2%).data(0).data0, 0) Then
    p2% = 0
    ty2 = 0
 End If
End If
If last_add_aid_point_for_two_line Mod 10 = 0 Then
   ReDim Preserve add_aid_point_for_two_line_(last_add_aid_point_for_two_line + 10) _
         As add_point_for_two_line_type
End If
last_add_aid_point_for_two_line = last_add_aid_point_for_two_line + 1
set_add_aid_point_for_two_line_ = last_add_aid_point_for_two_line
add_aid_point_for_two_line_(set_add_aid_point_for_two_line_).s_poi(0) = p1%
add_aid_point_for_two_line_(set_add_aid_point_for_two_line_).s_poi(1) = p2%
add_aid_point_for_two_line_(set_add_aid_point_for_two_line_).line_no(0) = l1%
add_aid_point_for_two_line_(set_add_aid_point_for_two_line_).line_no(1) = l2%
add_aid_point_for_two_line_(set_add_aid_point_for_two_line_).paral_or_verti(0) = ty1%
add_aid_point_for_two_line_(set_add_aid_point_for_two_line_).paral_or_verti(1) = ty2%

End Function

Public Sub simple_line_from_point_line(p%, ty%, l%)
Dim tl%
'画为代表形式，过p%点平行l%（垂直）的直线是
If ty = paral_ Then
 If is_point_in_paral_line(p%, l%, 0, tl%) Then
    p% = 0
    ty = 0
    l% = tl%
 End If
ElseIf ty = verti_ Then
  If is_point_in_verti_line(p%, l%, 0, tl%) Then
    p% = 0
    ty = 0
    l% = tl%
  End If
End If
End Sub



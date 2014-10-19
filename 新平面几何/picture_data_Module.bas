Attribute VB_Name = "picture_data_Module"
Option Explicit
Global Const move_point_no = -1
'******************************
Type T_line_type
visible As Byte
p(1) As POINTAPI
End Type
'Global l_p(3)
'Global tangent_line(10) As T_line_type
'******************************
Dim last_m_con_line As Integer
Dim last_m_aid_circle  As Integer
Dim last_m_aid_point  As Integer
Global last_m_aid_line   As Integer
'***************************************
 Global m_Con_lin() As line_type
'Global m_Circ() As circle_type
'Global m_aid_Circ() As circle_type '辅助圆,包括结论
Global m_lin() As line_type
Global aid_point(15) As Integer
Global red_line(15) As Integer
Global fill_color_line(7) As Integer
Global grey_line(7) As Integer
Dim last_aid_point As Integer
Dim last_m_line As Integer
Public Sub set_picture_data_init()
Dim i%
Erase m_poi
last_conditions.last_cond(1).point_no = 0
Erase m_lin
last_conditions.last_cond(1).line_no = 0
Erase m_Circ
last_conditions.last_cond(1).circle_no = 0
'Erase m_Con_lin
Erase m_aid_poi
last_conditions.last_cond(1).aid_point_no = 0
ReDim Preserve m_poi(100) As point_type
ReDim Preserve m_lin(100) As line_type
ReDim Preserve m_Circ(10) As circle_type
ReDim Preserve m_aid_poi(10) As point_type
last_m_aid_point = 10
last_m_line = 10
last_m_con_line = 0
last_m_circle = 10
last_m_aid_circle = 10
last_m_aid_line = 0
End Sub
Public Function set_point_coordinate(point_no%, coordinate0 As POINTAPI, is_change As Boolean) As Boolean
Dim i%
If point_no% = 0 Then
 For i% = 1 To last_conditions.last_cond(1).point_no
     If Abs(m_poi(i%).data(0).data0.coordinate.X - coordinate0.X) < 5 And _
         Abs(m_poi(i%).data(0).data0.coordinate.Y - coordinate0.Y) < 5 Then
          point_no% = i%
           set_point_coordinate = False
           Exit Function
     End If
 Next i%
ElseIf point_no% > 0 Then
     If Abs(m_poi(point_no%).data(0).data0.coordinate.X - coordinate0.X) < 5 And _
         Abs(m_poi(point_no%).data(0).data0.coordinate.Y - coordinate0.Y) < 5 Then
           set_point_coordinate = False
            Exit Function
     End If
End If
 set_point_coordinate = True
If point_no% = 0 Then
 If last_conditions.last_cond(1).point_no > 0 Then
    If C_display_wenti.is_last_wenti_complete = 1 Then
       Exit Function
    End If
 End If
 Call new_point_no(point_no%)
ElseIf point_no% = -1 Then
   point_no% = 0 '移动
    m_poi(0).data(0).data0.coordinate = coordinate0 '
     Exit Function
End If
  m_poi(point_no%).data(1) = m_poi(point_no%).data(0)
  m_poi(point_no%).data(0).data0.coordinate = coordinate0
  m_poi(point_no%).data(0).is_change = is_change
  If is_change Then
   Call change_line_(point_no%)
'   Call change_circle(point_no%)
  End If
    'Call change_line_(point_no%)
    'Call C_display_picture.PointCoordinateChange(point_no%)  '引起响应的圆变化
End Function
Public Sub move_point_inner_data(s As Byte, t As Byte)
Dim i%
For i% = 1 To last_conditions.last_cond(1).point_no
  m_poi(i%).data(t) = m_poi(i%).data(s)
Next i%
End Sub
Public Sub move_line_inner_data(s As Byte, t As Byte)
Dim i%
For i% = 1 To last_conditions.last_cond(1).line_no
  m_lin(i%).data(t) = m_lin(i%).data(s)
Next i%
End Sub
Public Sub move_point_data(s%, t%)
  m_poi(t%) = m_poi(s%)
End Sub
Public Sub move_line_data(s%, t%)
  m_lin(t%) = m_lin(s%)
End Sub
Public Sub set_point_name(ByVal point_no%, ByVal name$)
Dim i%
If point_no <= 100 And name$ >= "A" And name$ <= "Z" Then
Call new_point_no(point_no%)
 m_poi(point_no%).data(0).data0.name = name$
'***************************************************
   For i% = 1 To last_conditions.last_cond(1).point_no
   If (m_poi(i%).data(0).data0.name = "" Or _
        m_poi(i%).data(0).data0.name = global_icon_char) _
         And m_poi(i%).data(0).data0.visible > 0 Then
     choose_point = i%
        GoTo set_point_name_mark1
   End If
  Next i%
 choose_point = 0
set_point_name_mark1:
'******************************************
 Call C_display_picture.set_m_point_name(point_no%, name$)
 If operator = "re_name" Then
     Call C_display_wenti.re_name_for_point(point_no%, choose_point%)
      m_poi(point_no%).data(0).data0.name = name$
       Call C_display_picture.set_m_point_name(point_no%, name$)
 End If
ElseIf name$ <> m_poi(point_no%).data(0).data0.name Then
 m_poi(point_no%).data(0).data0.name = name$
  Call C_display_picture.set_m_point_name(point_no%, name$)
   Call C_display_wenti.re_name_for_point(point_no%, choose_point)
End If
If choose_point% > 0 And name <> "" Then
       Call C_display_picture.flash_point(choose_point)
End If
End Sub
Public Sub set_point_visible(ByVal point_no%, ByVal vi As Byte, is_change As Boolean)
Call new_point_no(point_no%)
 If m_poi(point_no%).data(0).data0.visible <> vi Then
  m_poi(point_no%).data(0).data0.visible = vi
   If vi = 1 And is_change Then
    m_poi(point_no%).data(0).is_change = True
    Call change_line_(point_no%)
   End If
 End If
    Call C_display_picture.set_m_point_visible(point_no%, vi)
End Sub
Public Sub set_line_visible(ByVal line_no%, ByVal vi As Byte)
If m_lin(line_no%).data(0).data0.type <> vi Then
   m_lin(line_no%).data(0).data0.type = vi
   Call C_display_picture.draw_line(line_no%, 0, 0)
End If
End Sub
Public Sub set_point_no_reduce(ByVal point_no%, ByVal no_reduce As Boolean)
Call new_point_no(point_no%)
End Sub
Public Sub set_point_in_circle(ByVal point_no%, ByVal c%)
Dim i%
If c% > 0 And point_no < 90 Then
Call new_point_no(point_no%)
For i% = 1 To m_poi(point_no%).data(0).in_circle(0)
    If m_poi(point_no%).data(0).in_circle(i%) = c% Then
      Exit Sub
    End If
Next i%
 m_poi(point_no%).data(0).in_circle(0) = i%
  m_poi(point_no%).data(0).in_circle(i%) = c%
End If
End Sub
Public Sub remove_circle_from_point0(ByVal point_no%, ByVal c%)
Dim i%, j%
Call new_point_no(point_no%)
For i% = 1 To m_poi(point_no%).data(0).in_circle(0)
    If m_poi(point_no%).data(0).in_circle(i%) = c% Then
       m_poi(point_no%).data(0).in_circle(0) = _
        m_poi(point_no%).data(0).in_circle(0) - 1
       For j% = i% To m_poi(point_no%).data(0).in_circle(0)
           m_poi(point_no%).data(0).in_circle(j%) = _
              m_poi(point_no%).data(0).in_circle(j% + 1)
       Next j%
      Exit Sub
    End If
Next i%
End Sub
Public Sub remove_circle_from_poi(ByVal c%)
Dim i%
For i% = 1 To last_conditions.last_cond(1).point_no
Call remove_circle_from_point0(i%, c%)
Next i%
End Sub

'Public Sub set_point_generate_by_line(ByVal point_no%, ByVal l1%, ByVal l2%)
'Call redim_poi(point_no%)
'If l1% > 0 Then
' If m_poi(point_no%).data(0).g_line(0) = 0 Then
'  m_poi(point_no%).data(0).g_line(0) = l1%
' Else
'  m_poi(point_no%).data(0).g_line(1) = l1%
' End If
' If m_poi(point_no%).data(0).depend_element(0).no = 0 Then
'  m_poi(point_no%).data(0).depend_element(0).ty = line_
'  m_poi(point_no%).data(0).depend_element(0).no = l1%
' Else
'  m_poi(point_no%).data(0).depend_element(1).ty = line_
'  m_poi(point_no%).data(0).depend_element(1).no = l1%
' End If
'End If
'If l2% > 0 Then
' If m_poi(point_no%).data(0).g_line(0) = 0 Then
'  m_poi(point_no%).data(0).g_line(0) = l2%
' Else
'  m_poi(point_no%).data(0).g_line(1) = l2%
' End If
' If m_poi(point_no%).data(0).depend_element(0).no = 0 Then
'  m_poi(point_no%).data(0).depend_element(0).ty = line_
'  m_poi(point_no%).data(0).depend_element(0).no = l2%
' Else
'  m_poi(point_no%).data(0).depend_element(1).ty = line_
'  m_poi(point_no%).data(0).depend_element(1).no = l2%
' End If
'End If
'End Sub
'Public Sub set_point_generate_by_circle(ByVal point_no%, ByVal c1%, ByVal c2%)
'Call redim_poi(point_no%)
'If c1% > 0 Then
' If m_poi(point_no%).data(0).g_circle(0) = 0 Then
'  m_poi(point_no%).data(0).g_circle(0) = c1%
' Else
'  m_poi(point_no%).data(0).g_circle(1) = c1%
' End If
'End If
'If c2% > 0 Then
' If m_poi(point_no%).data(0).g_circle(0) = 0 Then
'  m_poi(point_no%).data(0).g_circle(0) = c2%
' Else
'  m_poi(point_no%).data(0).g_circle(1) = c2%
' End If
'End If
'End Sub
Public Function set_point_in_line(ByVal point_no%, ByVal l%) As Boolean
Dim i%
If point_no% > 0 Then
For i% = 1 To m_poi(point_no%).data(0).in_line(0)
    If m_poi(point_no%).data(0).in_line(i%) = l% Then
      Exit Function
    End If
Next i%
 m_poi(point_no%).data(0).in_line(0) = i%
  m_poi(point_no%).data(0).in_line(i%) = l%
'   If point_no% = m_lin(l%).data(0).data0.poi(0) Or _
'       point_no% = m_lin(l%).data(0).data0.poi(1) Then
'       Call set_son_data(line_, l%, m_poi(point_no%).data(0).sons)
'   End If
    set_point_in_line = True
End If
End Function
Public Function set_line_in_verti(ByVal l1%, ByVal l2%, verti_no%, inter_point%) As Boolean
set_line_in_verti = set_line_in_verti0(l1%, l2%, verti_no%, inter_point%) And _
           set_line_in_verti0(l2%, l1%, verti_no%, inter_point%)
End Function
Private Function set_line_in_verti0(ByVal line_no%, ByVal l%, verti_no%, inter_point%) As Boolean
Dim i%, j%
Call redim_lin(line_no%)
For i% = 1 To m_lin(line_no%).data(0).in_verti(0).line_no
 If l% = m_lin(line_no%).data(0).in_verti(i%).line_no Then
    If inter_point% > 0 Then
     m_lin(line_no%).data(0).in_verti(i%).inter_point = inter_point%
    End If
     set_line_in_verti0 = False
  Exit Function
  '　已有垂直关系
 ElseIf l% < m_lin(line_no%).data(0).in_verti(i%).line_no Then
   For j% = m_lin(line_no%).data(0).in_verti(0).line_no To i% Step -1
     m_lin(line_no%).data(0).in_verti(j% + 1).line_no = m_lin(line_no%).data(0).in_verti(j%).line_no
   Next j%
     m_lin(line_no%).data(0).in_verti(0).line_no = m_lin(line_no%).data(0).in_verti(0).line_no + 1
     If l% > 0 Then
     m_lin(line_no%).data(0).in_verti(i%).line_no = l%
     End If
     If verti_no% > 0 Then
      m_lin(line_no%).data(0).in_verti(i%).verti_no = verti_no%
     End If
     If inter_point% > 0 Then
      m_lin(line_no%).data(0).in_verti(i%).inter_point = inter_point%
     End If
      set_line_in_verti0 = True
       Exit Function
 End If
 Next i%
   m_lin(line_no%).data(0).in_verti(0).line_no = m_lin(line_no%).data(0).in_verti(0).line_no + 1
    If l% > 0 Then
   m_lin(line_no%).data(0).in_verti(m_lin(line_no%).data(0).in_verti(0).line_no).line_no = _
           l%
   m_lin(line_no%).data(0).in_verti(m_lin(line_no%).data(0).in_verti(0).line_no).verti_no = _
           verti_no%
    m_lin(line_no%).data(0).in_verti(m_lin(line_no%).data(0).in_verti(0).line_no).inter_point = _
           inter_point%
 End If
  If verti_no% > 0 Then
   m_lin(line_no%).data(0).in_verti(m_lin(line_no%).data(0).in_verti(0).line_no).verti_no = _
           verti_no%
  End If
  If inter_point% > 0 Then
  m_lin(line_no%).data(0).in_verti(m_lin(line_no%).data(0).in_verti(0).line_no).inter_point = _
  inter_point%
  End If
End Function
Public Function set_line_in_paral(ByVal l1%, ByVal l2%, paral_no%) As Boolean
 set_line_in_paral = set_line_in_paral0(l1, l2%, paral_no%) And _
                     set_line_in_paral0(l2, l1%, paral_no%)
End Function
Private Function set_line_in_paral0(ByVal line_no%, ByVal l%, paral_no%) As Boolean
Dim i%, j%
Call redim_lin(line_no%)
For i% = 1 To m_lin(line_no%).data(0).in_paral(0).line_no
 If l% = m_lin(line_no%).data(0).in_paral(i%).line_no Then
    If paral_no% > 0 Then
       m_lin(line_no%).data(0).in_paral(i%).paral_no = paral_no%
    End If
   set_line_in_paral0 = False
  Exit Function
  '　已有垂直关系
 ElseIf l% < m_lin(line_no%).data(0).in_paral(i%).line_no Then
   For j% = m_lin(line_no%).data(0).in_paral(0).line_no To i% Step -1
     m_lin(line_no%).data(0).in_paral(j% + 1).line_no = m_lin(line_no%).data(0).in_paral(j%).line_no
   Next j%
     m_lin(line_no%).data(0).in_paral(0).line_no = m_lin(line_no%).data(0).in_paral(0).line_no + 1
     m_lin(line_no%).data(0).in_paral(i%).line_no = l%
     If paral_no% > 0 Then
        m_lin(line_no%).data(0).in_paral(i%).paral_no = paral_no%
     End If
        set_line_in_paral0 = True
       Exit Function
 End If
 Next i%
  m_lin(line_no%).data(0).in_paral(0).line_no = m_lin(line_no%).data(0).in_paral(0).line_no + 1
  m_lin(line_no%).data(0).in_paral(m_lin(line_no%).data(0).in_paral(0).line_no).line_no = l%
  m_lin(line_no%).data(0).in_paral(m_lin(line_no%).data(0).in_paral(0).line_no).paral_no = paral_no%
End Function

Public Function remove_line_from_point0(ByVal point_no%, ByVal l%) As Boolean
Dim i%, j%
Call new_point_no(point_no%)
For i% = 1 To m_poi(point_no%).data(0).in_line(0)
    If m_poi(point_no%).data(0).in_line(i%) = l% Then
       m_poi(point_no%).data(0).in_line(0) = _
           m_poi(point_no%).data(0).in_line(0) - 1
       For j% = i% To m_poi(point_no%).data(0).in_line(0)
           m_poi(point_no%).data(0).in_line(j%) = _
              m_poi(point_no%).data(0).in_line(j% + 1)
       Next j%
       remove_line_from_point0 = True
      Exit Function
    End If
Next i%
End Function
Public Sub remove_line_from_poi(ByVal point_no%, ByVal l%)
Dim i%
Call new_point_no(point_no%)
For i% = 1 To last_conditions.last_cond(1).point_no
Call remove_line_from_poi(i%, l%)
Next i%
End Sub
Public Sub replace_line_for_point(replace_no%, l%)
Dim i%
For i% = 1 To last_conditions.last_cond(1).point_no
 If remove_line_from_point0(i%, l%) Then
    Call set_point_in_line(i%, replace_no%)
 End If
Next i%
End Sub

Public Sub change_point_degree(ByVal point_no%, ByVal change_degree)
Dim t_d%
Call new_point_no(point_no%)
t_d% = m_poi(point_no%).data(0).degree + change_degree
If t_d% >= 0 Then
 m_poi(point_no%).data(0).degree = t_d%
End If
End Sub
Private Sub redim_lin(line_no%)
Dim i%
If last_m_line <= line_no% - 1 And last_conditions.last_cond(1).line_no Mod 100 = 0 Then
ReDim Preserve m_lin(last_m_line + 100) As line_type
For i% = last_m_line To last_m_line + 100
  m_lin(i%).data(0).other_no = i%
Next i%
last_m_line = last_m_line + 100
End If
End Sub
Public Function find_new_char() As String
Dim i%, k%
Dim ch As String
For k% = 65 To 90
For i% = 1 To last_conditions.last_cond(1).point_no
If m_poi(i%).data(0).data0.name = Chr(k%) Then
 GoTo find_new_char_mark1
End If
Next i%
For i% = 0 To last_aid_point_name
If add_point_name(i%) = ch Then
GoTo find_new_char_mark1
End If
Next i%
find_new_char = ch
find_new_char_mark1:
Next k%
End Function
Public Sub delete_line_from_poi(ByVal l%)
Dim i%
For i% = 1 To last_conditions.last_cond(1).point_no
      Call delete_line_from_point0(i%, l%)
Next i%
End Sub
Private Sub delete_line_from_point0(ByVal p%, ByVal l%)
Dim i%, j%
For i% = 1 To m_poi(p%).data(0).in_line(0)
    If m_poi(p%).data(0).in_line(i%) = l% Then
       m_poi(p%).data(0).in_line(0) = m_poi(p%).data(0).in_line(0) - 1
       For j% = i% To m_poi(p%).data(0).in_line(0)
        m_poi(p%).data(0).in_line(j%) = m_poi(p%).data(0).in_line(j% + 1)
       Next j%
       Exit Sub
    End If
Next i%
End Sub
Public Sub set_line_eangle_no(line_no%, eangle_no%)
Call redim_lin(line_no%)
 m_lin(line_no%).data(0).eangle_no = eangle_no%
End Sub
Public Sub set_line_direction(line_no%, dir As Integer)
Call redim_lin(line_no%)
 m_lin(line_no%).data(0).data0.in_point(10) = dir
End Sub

Public Sub set_line_other_no(line_no%, other_no%)
m_lin(line_no%).data(0).other_no = other_no%
End Sub
'Public Sub set_line_no_reduce(line_no%, no_reduce As Boolean)
'm_lin(line_no%).data(0).no_reduce = no_reduce
'End Sub
Public Sub set_line_cond_data(line_no%, cond_data As condition_data_type)
m_lin(line_no%).data(0).cond_data = cond_data
End Sub


Public Function simple_dbase_for_line(replace_no%, no1%, no2%, ByVal new_p%, re As record_data_type) As Byte
Dim i%, j%, k%, l%, m%, n%, no%
Dim tn%, tn_%, tA%
Dim t_n() As Integer
Dim last_tn%
Dim ts As String
Dim tp(3) As Integer
Dim ty As Boolean
Dim n_(8) As Integer
Dim n1_(8) As Integer
Dim t_l As two_line_type
Dim A As angle_data_type
Dim item_0 As item0_data_type
Dim tan_l As tangent_line_data_type
Dim temp_record As total_record_type
Dim T_a(1) As total_angle_data_type
Dim t_a_v As angle3_value_data0_type
Dim combine_line(2, 20) As Integer
Dim last_combine_line As Integer
Dim is_remove As Boolean
Dim is_exist As Boolean
'On Error GoTo simple_dbase_for_line
last_subs_angle3_value = 0
Erase subs_angle3_value
If no1% = replace_no% Then
   no1% = 0
ElseIf no2% = replace_no% Then
   no2% = 0
End If
'For i% = 1 To last_conditions.last_cond(1).point_no
'    Call replace_line_for_point(replace_no%, no1%)
'    Call replace_line_for_point(replace_no%, no2%)
'Next i%
If m_lin(no1%).data(0).data0.visible >= m_lin(no2%).data(0).data0.visible Then '显示replace_no%
 If m_lin(no1%).data(0).data0.visible > m_lin(replace_no%).data(0).data0.visible Then
  Call set_line_visible(replace_no%, m_lin(no1%).data(0).data0.visible)
 End If
Else
 If m_lin(no2%).data(0).data0.visible > m_lin(replace_no%).data(0).data0.visible Then
  Call set_line_visible(replace_no%, m_lin(no2%).data(0).data0.visible)
 End If
End If
'更改数据
For k% = 1 To last_conditions.last_cond(1).polygon4_no
 For j% = 0 To 3
  If Dpolygon4(k%).data(0).line_no(j%) = no1% Or Dpolygon4(k%).data(0).line_no(j%) = no2% Then
   Dpolygon4(k%).data(0).line_no(j%) = replace_no%
  End If
 Next j%
Next k%
'***********************************************
For i% = 1 To last_conditions.last_cond(1).line_no
  If i% <> no1% And i% <> no2% And i% <> replace_no% Then
    n1_(0) = total_angle_no(replace_no%, i%)
    n1_(1) = total_angle_no(no1%, i%)
    n1_(2) = total_angle_no(no2%, i%)
     Call combine_two_total_angle(n1_(0), n1_(1), i%, replace_no%, no1%)
     Call combine_two_total_angle(n1_(0), n1_(2), i%, replace_no%, no2%)
     For j% = 1 To 4
       If T_angle(tn%).data(0).angle_no(j%).no > 0 Then
         Call add_record_to_record(re.data0.condition_data, angle(T_angle(tn%).data(0).angle_no(j%).no).data(0).cond_data)
       End If
     Next j%
  End If
Next i%
 n1_(0) = no1%
  n1_(1) = no2%
   n1_(2) = replace_no%
   
For m% = 0 To 1
 For n% = 0 To 1
  If n1_(n%) > 0 Then
    T_a(1).line_no(0) = i%
    T_a(1).line_no(1) = replace_no%
     If search_for_total_angle(T_a(1), n_(1)) Then
      Call retire_data(total_angle_, n_(1))
     End If
'*************************************************
   For i% = 1 To last_conditions.last_cond(1).line_no 'n1_(n%) - 1
    If i% < n1_(n%) Then
     T_a(0).line_no(0) = i%
     T_a(0).line_no(1) = n1_(0)
     T_a(1).line_no(0) = i%
     T_a(1).line_no(1) = replace_no%
    ElseIf i% > n1_(n%) And i% < n1_(2) Then
     T_a(0).line_no(0) = n1_(0)
     T_a(0).line_no(1) = i%
     T_a(1).line_no(0) = i%
     T_a(1).line_no(1) = replace_no%
    Else
     T_a(0).line_no(0) = n1_(0)
     T_a(0).line_no(1) = i%
     T_a(1).line_no(0) = replace_no%
     T_a(1).line_no(1) = i%
    End If
    If search_for_total_angle(T_a(0), n_(0)) Then
     'If search_for_total_angle(T_a(1), n_(1)) = False Then
      n_(1) = set_total_angle(T_a(1))
      If T_angle(n_(1)).data(0).is_used_no = 0 Then
        If T_angle(n_(0)).data(0).line_no(0) = T_angle(n_(1)).data(0).line_no(0) _
            Or T_angle(n_(0)).data(0).line_no(1) = T_angle(n_(1)).data(0).line_no(1) Then
             
        Else
        End If
      End If
     'End If
     Call retire_data(total_angle_, n_(0))
    End If
    'If n_(1) > 0 And n_(0) > 0 Then
    'End If
   Next i%
'For k% = n_(0) + 1 To n_(1)
'  tn% = T_angle(k%).data(0).index(m%)
'   last_tn% = last_tn% + 1
'   ReDim Preserve t_n(last_tn%) As Integer
'   t_n(last_tn%) = tn%
'Next k%
'For k% = 1 To last_tn%
' tn% = t_n(k%) '与no1%,no2%有关的角
'  If T_angle(tn%).data(0).line_no((m% + 1) Mod 2) = n1_((n% + 1) Mod 2) Or _
'                    T_angle(tn%).data(0).line_no((m% + 1) Mod 2) = n1_((n% + 2) Mod 2) Then '另一边也等于替换边,平角
'    tA% = T_angle(tn%).data(0).is_used_no '
'     If tA% >= 0 Then
'        tA% = T_angle(tn%).data(0).angle_no(tA%).no '全角的代表角
'          Call is_point_in_line3(angle(tA%).data(0).poi(0), m_lin(replace_no%).data(0).data0, n_(0))
'          Call is_point_in_line3(angle(tA%).data(0).poi(1), m_lin(replace_no%).data(0).data0, n_(1))
'          Call is_point_in_line3(angle(tA%).data(0).poi(2), m_lin(replace_no%).data(0).data0, n_(2))
'          ts = ""
'          If (n_(1) - n_(0)) * (n_(1) - n_(2)) > 0 Then '被替换的角的大小
'              ts = "0"
'              ty = 0
'          Else
'              ts = "180"
'              ty = 1
'          End If
'          temp_record.record_data = re
'        simple_dbase_for_line = simple_dbase_for_angle(tA%, 0, ty, temp_record) 'simple_three_angle_value(tA%, ts, T_angle(tn%).data(0).is_used_no, re, 0)
'         If simple_dbase_for_line > 1 Then
'           Exit Function
'         End If
'     End If
'    If n% < 2 Then
'    If T_angle(tn%).data(0).angle_no(0).no > 0 Then
'     Call remove_record(angle_, T_angle(tn%).data(0).angle_no(0).no, 0)
'    End If
'    If T_angle(tn%).data(0).angle_no(1).no > 0 Then
'    Call remove_record(angle_, T_angle(tn%).data(0).angle_no(1).no, 0)
'    End If
'    If T_angle(tn%).data(0).angle_no(2).no > 0 Then
'    Call remove_record(angle_, T_angle(tn%).data(0).angle_no(2).no, 0)
'    End If
'    If T_angle(tn%).data(0).angle_no(3).no > 0 Then
'    Call remove_record(angle_, T_angle(tn%).data(0).angle_no(3).no, 0)
'    End If
'    Call remove_record(total_angle_, tn%, 0)
'    End If
'  Else
'    tA% = T_angle(tn%).data(0).is_used_no '
'     t_A(0) = T_angle(tn%).data(0)
'     t_A(1) = T_angle(tn%).data(0)
'     t_A(0).line_no(m%) = replace_no%
'     no% = 0
'  If search_for_total_angle(t_A(0), no%, 0, 0) Then   '相应的全角
'     If no% > 0 Then
'     t_A(0) = T_angle(no%).data(0)
'***************************
'     If tA% = -1 Then '无交点全角
'     ElseIf t_A(0).is_used_no >= 0 Then
'        tA% = T_angle(tn%).data(0).angle_no(tA%).no '全角的代表角
'         ts = ""
         '***************************************
'         temp_record.record_data = re
'         For i% = 0 To 3
'            If t_A(1).angle_no(i%).no > 0 And t_A(0).angle_no(i%).no Then
'             angle(t_A(1).angle_no(i%).no).data(0).other_no = t_A(0).angle_no(i%).no
'             Call remove_record(angle_, t_A(1).angle_no(0).no, 0)
'            End If
'         Next i%
'         If Abs(t_A(1).is_used_no - t_A(0).is_used_no) Mod 2 = 0 Then
'           simple_dbase_for_line = simple_dbase_for_angle( _
'             t_A(1).angle_no(t_A(1).is_used_no).no, _
'              t_A(0).angle_no(t_A(0).is_used_no).no, 0, temp_record)
'           If simple_dbase_for_line > 1 Then
'              Exit Function
'           End If
'         Else
'           simple_dbase_for_line = simple_dbase_for_angle( _
'             t_A(1).angle_no(t_A(1).is_used_no).no, _
'              t_A(0).angle_no(t_A(0).is_used_no).no, 1, temp_record)
'           If simple_dbase_for_line > 1 Then
'              Exit Function
'           End If
'         End If
'     End If
'    End If
'    Else
'     no% = 0
'     Call set_total_angle(t_A(0), 0, no%, False)
'    End If
   '********************************************
'   End If
  '**************************************
'   '  t_A(0) = T_angle(tn%).data(0)
'   '  t_A(1) = T_angle(tn%).data(0)
'   '  t_A(0).line_no(m%) = replace_no%
'   '  no% = 0
'   '  If search_for_total_angle(t_A(0), no%, 0, 0) Then '相应的全角
'        't_A(1) = T_angle(tn%).data(0)
'         If tn% <> no% And no% > 0 Then
'         t_A(0) = T_angle(no%).data(0)
'       For i% = 0 To 3
'       If t_A(0).angle_no(i%).no > 0 And t_A(1).angle_no(i%).no > 0 Then
'        If t_A(0).angle_no(i%).no <> t_A(1).angle_no(i%).no Then '相应的真角不等
         '******************************************************************************
'          temp_record.record_data.data0.condition_data.condition_no = 0
'            simple_dbase_for_line = set_total_equal_triangle_from_eangle(t_A(1).angle_no(i%).no, _
'             t_A(0).angle_no(i%).no, temp_record, new_p%, 0, 0, 0, 0, 0, 1)
'              If simple_dbase_for_line > 1 Then
'               Exit Function
'              End If
       '**************************************************************************************
'         For l% = last_conditions.last_cond(0).triangle_no + 1 To last_conditions.last_cond(1).triangle_no
'              tA% = triangle(l%).data(0).index.i(0)
'           For j% = 0 To 2
'            If triangle(tA%).data(0).angle(j%) = t_A(1).angle_no(i%).no Then
'              If ts = "" Then
'               triangle(tA%).data(0).angle(j%) = t_A(0).angle_no(i%).no
'              Else
'               Call remove_record(triangle_, tA%, 0)
'              End If
'            End If
'           Next j%
'         Next l%
'            GoTo simple_dbase_for_line_next1
'          End If
'       End If'
'       Next i%
'      End If
'simple_dbase_for_line_next1:
'     Call remove_record(total_angle_, tn%, 0)
'' End If
'Next k%
End If
Next n%
Next m%
simple_dbase_for_line = simple_dbase_for_angle_(re)
'*********************************************************
For k% = 0 To last_conclusion - 1
 If conclusion_data(k%).ty = verti_mid_line_ And conclusion_data(k%).no(0) = 0 Then
    If con_verti_mid_line(k%).data(0).data0.line_no(0) = no1% Or _
             con_verti_mid_line(k%).data(0).data0.line_no(0) = no2 Then
              con_verti_mid_line(k%).data(0).data0.line_no(0) = replace_no%
               Call line_number0(con_verti_mid_line(k%).data(0).data0.poi(0), _
                 con_verti_mid_line(k%).data(0).data0.poi(2), _
                  con_verti_mid_line(k%).data(0).data0.n(0), _
                   con_verti_mid_line(k%).data(0).data0.n(2))
               Call is_point_in_line3(con_verti_mid_line(k%).data(0).data0.poi(1), _
                     m_lin(replace_no%).data(0).data0, con_verti_mid_line(k%).data(0).data0.n(1))
    ElseIf con_verti_mid_line(k%).data(0).data0.line_no(1) = no1% Or _
              con_verti_mid_line(k%).data(0).data0.line_no(1) = no2 Then
              con_verti_mid_line(k%).data(0).data0.line_no(1) = replace_no%
    End If
 End If
Next k%
For k% = 1 + last_conditions.last_cond(0).verti_mid_line_no To last_conditions.last_cond(1).verti_mid_line_no
i% = verti_mid_line(k%).data(0).record.data1.index.i(0)
If verti_mid_line(i%).record_.no_reduce < 255 Then
If verti_mid_line(i%).data(0).data0.line_no(0) = no1% Or _
    verti_mid_line(i%).data(0).data0.line_no(0) = no2% Or _
     verti_mid_line(i%).data(0).data0.line_no(0) = replace_no% Then
tn% = 0
n_(0) = -5000
If is_verti_mid_line( _
        verti_mid_line(i%).data(0).data0.poi(0), verti_mid_line(i%).data(0).data0.poi(1), _
         verti_mid_line(i%).data(0).data0.poi(2), verti_mid_line(i%).data(0).data0.line_no(0), _
          tn%, n_(0), n_(1), verti_mid_line_data0) Then
           If tn% <> i% Then
            Call remove_record(verti_mid_line_, i%, 0)
           End If
Else
 Call search_for_verti_mid_line(verti_mid_line(i%).data(0).data0, n1_(0), 0, 1)
 Call search_for_verti_mid_line(verti_mid_line(i%).data(0).data0, n1_(1), 1, 1)
  verti_mid_line(i%).data(0).data0 = verti_mid_line_data0
   For j% = 0 To 1
    If n1_(j%) < n_(j%) Then
     For l% = n1_(j%) + 1 To n_(j%) - 1
      verti_mid_line(l%).data(0).record.data1.index.i(j%) = verti_mid_line(l% + 1).data(0).record.data1.index.i(j%)
     Next l%
     verti_mid_line(n_(j%)).data(0).record.data1.index.i(j%) = i%
    ElseIf n1_(j%) > n_(j%) Then
     For l% = n1_(j%) + 1 To n_(j%) + 2 Step -1
      verti_mid_line(l%).data(0).record.data1.index.i(j%) = verti_mid_line(l% - 1).data(0).record.data1.index.i(j%)
     Next l%
     verti_mid_line(n_(j%) + 1).data(0).record.data1.index.i(j%) = i%
    End If
   Next j%
End If
End If
End If
Next k%
'*******************
For k% = 0 To last_conclusion - 1
 If conclusion_data(k%).ty = eline_ And conclusion_data(k%).no(0) = 0 Then
   If con_eline(k%).data(0).data0.line_no(0) = no1% Or con_eline(k%).data(0).data0.line_no(0) = no2% Or _
        con_eline(k%).data(0).data0.line_no(0) = replace_no% Or con_eline(k%).data(0).data0.line_no(1) = no1% Or _
         con_eline(k%).data(0).data0.line_no(1) = no2% Or con_eline(k%).data(0).data0.line_no(1) = replace_no% Then
    Call is_equal_dline(con_eline(k%).data(0).data0.poi(0), con_eline(k%).data(0).data0.poi(1), _
                        con_eline(k%).data(0).data0.poi(2), con_eline(k%).data(0).data0.poi(3), _
                         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, con_eline(k%).data(0).data0, 0, 0, 0, "", _
                          temp_record.record_data.data0.condition_data)
   End If
 End If
Next k%
For k% = 1 + last_conditions.last_cond(0).eline_no To last_conditions.last_cond(1).eline_no
i% = Deline(k%).data(0).record.data1.index.i(0)
If Deline(i%).record_.no_reduce < 255 Then
If Deline(i%).data(0).data0.line_no(0) = no1% Or Deline(i%).data(0).data0.line_no(0) = no2% Or _
    Deline(i%).data(0).data0.line_no(0) = replace_no% Or Deline(i%).data(0).data0.line_no(1) = no1% Or _
     Deline(i%).data(0).data0.line_no(1) = no2% Or Deline(i%).data(0).data0.line_no(1) = replace_no% Then
tn% = 0
record_0.data0.condition_data.condition_no = 0 'record0
n_(0) = -5000
ty = is_equal_dline( _
        Deline(i%).data(0).data0.poi(0), Deline(i%).data(0).data0.poi(1), _
         Deline(i%).data(0).data0.poi(2), Deline(i%).data(0).data0.poi(3), _
          0, 0, 0, 0, 0, 0, tn%, n_(0), n_(1), n_(2), n_(3), _
             eline_data0, 0, 0, 0, "", record_0.data0.condition_data)
 If eline_data0.line_no(0) = eline_data0.line_no(1) And eline_data0.n(1) = eline_data0.n(2) Then
  temp_record.record_data = Deline(i%).data(0).record
       Call set_mid_point(eline_data0.poi(0), eline_data0.poi(1), eline_data0.poi(3), _
         eline_data0.n(0), eline_data0.n(1), eline_data0.n(3), eline_data0.line_no(0), _
          0, temp_record, 0, 0, 0, 0, 0)
           Call remove_record(eline_, i%, 0)
 Else
  If ty Then
   If i% <> tn% Then
    If Deline(i%).data(0).record.data0.condition_data.level < Deline(tn%).data(0).record.data0.condition_data.level Then
     Call remove_record(eline_, tn%, 0)
    Else
     Call remove_record(eline_, i%, 0)
    End If
   Else
    Deline(i%).data(0).data0 = eline_data0
     Call search_for_eline(Deline(i%).data(0).data0, 2, n1_(2), 1)
     Call search_for_eline(Deline(i%).data(0).data0, 3, n1_(3), 1)
    For j% = 2 To 3
     If n_(j%) < n1_(j%) Then
      For l% = n1_(j%) + 1 To n_(j%) + 2 Step -1
       Deline(l%).data(0).record.data1.index.i(j%) = Deline(l% - 1).data(0).record.data1.index.i(j%)
      Next l%
      Deline(n_(j%) + 1).data(0).record.data1.index.i(j%) = i%
      ElseIf n_(j%) > n1_(j%) Then
       For l% = n1_(j%) + 1 To n_(j%) - 1
      Deline(l%).data(0).record.data1.index.i(j%) = Deline(l% + 1).data(0).record.data1.index.i(j%)
      Next l%
      Deline(n_(j%)).data(0).record.data1.index.i(j%) = i%
     End If
   Next j%
    End If
  Else
   Call search_for_eline(Deline(i%).data(0).data0, 0, n1_(0), 1)
   Call search_for_eline(Deline(i%).data(0).data0, 1, n1_(1), 1)
   Call search_for_eline(Deline(i%).data(0).data0, 2, n1_(2), 1)
   Call search_for_eline(Deline(i%).data(0).data0, 3, n1_(3), 1)
   'Call search_for_eline(Deline(i%).data(0).data0, 4, n1_(4), 1)
   Deline(i%).data(0).data0 = eline_data0
   For j% = 0 To 3
    If n_(j%) < n1_(j%) Then
     For l% = n1_(j%) + 1 To n_(j%) + 2 Step -1
      Deline(l%).data(0).record.data1.index.i(j%) = Deline(l% - 1).data(0).record.data1.index.i(j%)
     Next l%
    Deline(n_(j%) + 1).data(0).record.data1.index.i(j%) = i%
    ElseIf n_(j%) > n1_(j%) Then
     For l% = n1_(j%) + 1 To n_(j%) - 1
      Deline(l%).data(0).record.data1.index.i(j%) = Deline(l% + 1).data(0).record.data1.index.i(j%)
     Next l%
    Deline(n_(j%)).data(0).record.data1.index.i(j%) = i%
    End If
   Next j%
   End If
End If
End If
End If
Next k%
'*************
For k% = 0 To last_conclusion - 1
 If conclusion_data(k%).ty = line_value_ And conclusion_data(k%).no(0) = 0 Then
    If con_line_value(k%).data(0).data0.line_no = no1% Or con_line_value(k%).data(0).data0.line_no = no2% Or _
       con_line_value(k%).data(0).data0.line_no = replace_no% Then
       Call is_line_value(con_line_value(k%).data(0).data0.poi(0), con_line_value(k%).data(0).data0.poi(1), _
                     0, 0, 0, con_line_value(k%).data(0).data0.value, 0, 0, 0, 0, 0, _
                       con_line_value(k%).data(0).data0)
    End If
 End If
Next k%

For k% = 1 + last_conditions.last_cond(0).line_value_no To last_conditions.last_cond(1).line_value_no
i% = line_value(k%).data(0).record.data1.index.i(0)
If line_value(i%).record_.no_reduce < 255 Then
If line_value(i%).data(0).data0.line_no = no1% Or line_value(i%).data(0).data0.line_no = no2% Or _
     line_value(i%).data(0).data0.line_no = replace_no% Then
     tn% = 0
record_0.data0.condition_data.condition_no = 0 ' record0
'n_(0) = -5000
If is_line_value(line_value(i%).data(0).data0.poi(0), line_value(i%).data(0).data0.poi(1), _
     0, 0, 0, line_value(i%).data(0).data0.value, tn%, n_(0), n_(1), n_(2), _
      n_(3), line_value_data0) = 1 Then
      line_value_data0.value_ = line_value(i%).data(0).data0.value_
      line_value_data0.squar_value = line_value(i%).data(0).data0.squar_value
If tn% <> i% Then
  If line_value(i%).data(0).record.data0.condition_data.level < line_value(tn%).data(0).record.data0.condition_data.level Then
   Call remove_record(line_value_, tn%, 0)
  Else
   Call remove_record(line_value_, i%, 0)
  End If
  line_value(i%).data(0).data0 = line_value_data0
Else
 line_value(i%).data(0).data0 = line_value_data0
End If
Else
Call search_for_line_value(line_value(i%).data(0).data0, 0, n1_(0), 1)
Call search_for_line_value(line_value(i%).data(0).data0, 1, n1_(1), 1)
Call search_for_line_value(line_value(i%).data(0).data0, 2, n1_(2), 1)
Call search_for_line_value(line_value(i%).data(0).data0, 3, n1_(3), 1)
line_value(i%).data(0).data0 = line_value_data0
For j% = 0 To 3
 If n1_(j%) < n_(j%) Then
  For l% = n1_(j%) + 1 To n_(j%) + 2 Step -1
   line_value(l%).data(0).record.data1.index.i(j%) = line_value(l% - 1).data(0).record.data1.index.i(j%)
  Next l%
   line_value(n_(j%) + 1).data(0).record.data1.index.i(j%) = i%
 ElseIf n1_(j%) > n_(j%) Then
  For l% = n1_(j%) + 1 To n_(j%) - 1
   line_value(l%).data(0).record.data1.index.i(j%) = line_value(l% + 1).data(0).record.data1.index.i(j%)
  Next l%
   line_value(n_(j%)).data(0).record.data1.index.i(j%) = i%
 End If
Next j%
End If
End If
End If
Next k%
'***********************
For k% = 0 To last_conclusion - 1
 If conclusion_data(k%).ty = two_line_value_ And conclusion_data(k%).no(0) = 0 Then
    If con_two_line_value(k%).data(0).line_no(0) = no1% Or con_two_line_value(k%).data(0).line_no(0) = no2% Or _
       con_two_line_value(k%).data(0).line_no(0) = replace_no% Or con_two_line_value(k%).data(0).line_no(1) = no1% Or _
       con_two_line_value(k%).data(0).line_no(1) = no2% Or con_two_line_value(k%).data(0).line_no(1) = replace_no% Then
       Call is_two_line_value(con_two_line_value(k%).data(0).poi(0), con_two_line_value(k%).data(0).poi(1), _
                                 con_two_line_value(k%).data(0).poi(2), con_two_line_value(k%).data(0).poi(3), _
                                   0, 0, 0, 0, 0, 0, con_two_line_value(k%).data(0).para(0), _
                                  con_two_line_value(k%).data(0).para(1), con_two_line_value(k%).data(0).value, _
                         0, 0, 0, 0, 0, con_two_line_value(k%).data(0), 0, temp_record.record_data.data0.condition_data)
                                  
    End If
 End If
Next k%
For k% = 1 + last_conditions.last_cond(0).two_line_value_no To last_conditions.last_cond(1).two_line_value_no
 i% = two_line_value(k%).data(0).record.data1.index.i(0)
If two_line_value(i%).record_.no_reduce < 255 Then
If two_line_value(i%).data(0).data0.line_no(0) = no1% Or two_line_value(i%).data(0).data0.line_no(0) = no2% Or _
     two_line_value(i%).data(0).data0.line_no(0) = replace_no% Or two_line_value(i%).data(0).data0.line_no(1) = no1% Or _
      two_line_value(i%).data(0).data0.line_no(1) = no2% Or two_line_value(i%).data(0).data0.line_no(1) = replace_no% Then
tn% = 0
record_0.data0.condition_data.condition_no = 0 'record0
n_(0) = -5000
If is_two_line_value(two_line_value(i%).data(0).data0.poi(0), two_line_value(i%).data(0).data0.poi(1), _
    two_line_value(i%).data(0).data0.poi(2), two_line_value(i%).data(0).data0.poi(3), 0, 0, 0, 0, 0, 0, _
     two_line_value(i%).data(0).data0.para(0), two_line_value(i%).data(0).data0.para(1), _
      two_line_value(i%).data(0).data0.value, tn%, n_(0), n_(1), n_(2), n_(3), _
        two_line_value_data0, 0, record_0.data0.condition_data) = 1 Then
If tn% <> i% Then
 If two_line_value(i%).data(0).record.data0.condition_data.level < two_line_value(tn%).data(0).record.data0.condition_data.level Then
  Call remove_record(two_line_value_, tn%, 0)
 Else
  Call remove_record(two_line_value_, i%, 0)
 End If
  two_line_value(i%).data(0).data0 = two_line_value_data0
Else
 two_line_value(i%).data(0).data0 = two_line_value_data0
End If
Else
 Call search_for_two_line_value(two_line_value(i%).data(0).data0, 0, n1_(0), 1)
 Call search_for_two_line_value(two_line_value(i%).data(0).data0, 1, n1_(1), 1)
 Call search_for_two_line_value(two_line_value(i%).data(0).data0, 2, n1_(2), 1)
 Call search_for_two_line_value(two_line_value(i%).data(0).data0, 3, n1_(3), 1)
 two_line_value(i%).data(0).data0 = two_line_value_data0
 For j% = 0 To 3
  If n1_(j%) < n_(j%) Then
   For l% = n1_(j%) + 1 To n_(j%) - 1
    two_line_value(l%).data(0).record.data1.index.i(j%) = two_line_value(l% + 1).data(0).record.data1.index.i(j%)
   Next l%
    two_line_value(n_(j%)).data(0).record.data1.index.i(j%) = i%
  ElseIf n1_(j%) > n_(j%) Then
   For l% = n1_(j%) + 1 To n_(j%) + 2 Step -1
    two_line_value(l%).data(0).record.data1.index.i(j%) = two_line_value(l% - 1).data(0).record.data1.index.i(j%)
   Next l%
    two_line_value(n_(j%) + 1).data(0).record.data1.index.i(j%) = i%
 End If
 Next j%
End If
End If
End If
Next k%
'****************************************
For k% = 0 To last_conclusion - 1
 If conclusion_data(k%).ty = line3_value_ And conclusion_data(k%).no(0) = 0 Then
    If con_line3_value(k%).data(0).line_no(0) = no1% Or con_line3_value(k%).data(0).line_no(0) = no2% Or _
       con_line3_value(k%).data(0).line_no(0) = replace_no% Or con_line3_value(k%).data(0).line_no(1) = no1% Or _
       con_line3_value(k%).data(0).line_no(1) = no2% Or con_line3_value(k%).data(0).line_no(1) = replace_no% Or _
       con_line3_value(k%).data(0).line_no(2) = no1% Or con_line3_value(k%).data(0).line_no(2) = no2% Or _
       con_line3_value(k%).data(0).line_no(2) = replace_no% Then
       Call is_three_line_value(con_line3_value(k%).data(0).poi(0), con_line3_value(k%).data(0).poi(1), _
                                 con_line3_value(k%).data(0).poi(2), con_line3_value(k%).data(0).poi(3), _
                                  con_line3_value(k%).data(0).poi(4), con_line3_value(k%).data(0).poi(5), _
                    0, 0, 0, 0, 0, 0, 0, 0, 0, con_line3_value(k%).data(0).para(0), _
                     con_line3_value(k%).data(0).para(1), con_line3_value(k%).data(0).para(2), _
                      con_line3_value(k%).data(0).value, 0, 0, 0, 0, 0, 0, 0, con_line3_value(k%).data(0), 0, _
                       temp_record.record_data.data0.condition_data, 0)
                                  
    End If
 End If
Next k%
For k% = 1 + last_conditions.last_cond(0).line3_value_no To last_conditions.last_cond(1).line3_value_no
i% = line3_value(k%).data(0).record.data1.index.i(0)
If line3_value(i%).record_.no_reduce < 255 Then
If line3_value(i%).data(0).data0.line_no(0) = no1% Or line3_value(i%).data(0).data0.line_no(0) = no2% Or _
     line3_value(i%).data(0).data0.line_no(1) = no1% Or line3_value(i%).data(0).data0.line_no(1) = no2% Or _
      line3_value(i%).data(0).data0.line_no(2) = no1% Or line3_value(i%).data(0).data0.line_no(2) = no2% Or _
       line3_value(i%).data(0).data0.line_no(0) = replace_no% Or _
        line3_value(i%).data(0).data0.line_no(1) = replace_no% Or _
         line3_value(i%).data(0).data0.line_no(2) = replace_no% Then
 tn% = 0
  record_0.data0.condition_data.condition_no = 0 'record0
   n_(0) = -5000
 If is_three_line_value(line3_value(i%).data(0).data0.poi(0), line3_value(i%).data(0).data0.poi(1), _
    line3_value(i%).data(0).data0.poi(2), line3_value(i%).data(0).data0.poi(3), line3_value(i%).data(0).data0.poi(4), _
     line3_value(i%).data(0).data0.poi(5), 0, 0, 0, 0, 0, 0, 0, 0, 0, _
      line3_value(i%).data(0).data0.para(0), line3_value(i%).data(0).data0.para(1), _
       line3_value(i%).data(0).data0.para(2), line3_value(i%).data(0).data0.value, tn%, n_(0), n_(1), _
        n_(2), n_(3), n_(4), n_(5), line3_value_data0, 0, _
         record_0.data0.condition_data, 0) = 1 Then
If i% <> tn% Then
 If line3_value(i%).data(0).record.data0.condition_data.level < line3_value(tn%).data(0).record.data0.condition_data.level Then
 Call remove_record(line3_value_, tn%, 0)
 Else
 Call remove_record(line3_value_, i%, 0)
 End If
Else
 line3_value(i%).data(0).data0 = line3_value_data0
End If
Else
 Call search_for_line3_value(line3_value(i%).data(0).data0, 0, n1_(0), 1)
 Call search_for_line3_value(line3_value(i%).data(0).data0, 1, n1_(1), 1)
 Call search_for_line3_value(line3_value(i%).data(0).data0, 2, n1_(2), 1)
 Call search_for_line3_value(line3_value(i%).data(0).data0, 3, n1_(3), 1)
 Call search_for_line3_value(line3_value(i%).data(0).data0, 4, n1_(4), 1)
 Call search_for_line3_value(line3_value(i%).data(0).data0, 5, n1_(5), 1)
 line3_value(i%).data(0).data0 = line3_value_data0
 For j% = 0 To 5
  If n_(j%) < n1_(j%) Then
   For l% = n1_(j%) + 1 To n_(j%) + 2 Step -1
    line3_value(l%).data(0).record.data1.index.i(j%) = line3_value(l% - 1).data(0).record.data1.index.i(j%)
   Next l%
    line3_value(n_(j%) + 1).data(0).record.data1.index.i(j%) = i%
  ElseIf n_(j%) > n1_(j%) Then
   For l% = n1_(j%) + 1 To n_(j%) - 1
    line3_value(l%).data(0).record.data1.index.i(j%) = line3_value(l% + 1).data(0).record.data1.index.i(j%)
   Next l%
    line3_value(n_(j%)).data(0).record.data1.index.i(j%) = i%
  End If
 Next j%
End If
End If
End If
Next k%
For k% = 0 To last_conclusion - 1
 If conclusion_data(k%).ty = paral_ And conclusion_data(k%).no(0) = 0 Then
    If con_paral(k%).data(0).line_no(0) = no1% Or con_paral(k%).data(0).line_no(0) = no2% Then
        con_paral(k%).data(0).line_no(0) = replace_no%
    ElseIf con_paral(k%).data(0).line_no(1) = no1% Or con_paral(k%).data(0).line_no(1) = no2% Then
        con_paral(k%).data(0).line_no(1) = replace_no%
    End If
    If con_paral(k%).data(0).line_no(0) > con_paral(k%).data(0).line_no(1) Then
     Call exchange_two_integer(con_paral(k%).data(0).line_no(0), con_paral(k%).data(0).line_no(1))
    End If
 End If
Next k%
For k% = 1 + last_conditions.last_cond(0).paral_no To last_conditions.last_cond(1).paral_no
 i% = Dparal(k%).data(0).data0.record.data1.index.i(0)
 If Dparal(i%).record_.no_reduce < 255 Then
 t_l = Dparal(i%).data(0).data0
 If t_l.line_no(0) = no1% Or t_l.line_no(0) = no2% Or _
       t_l.line_no(1) = no1% Or t_l.line_no(1) = no2% Then
 If t_l.line_no(0) = no1% Or t_l.line_no(0) = no2% Then
    t_l.line_no(0) = replace_no%
 End If
 If t_l.line_no(1) = no1% Or t_l.line_no(1) = no2% Then
    t_l.line_no(1) = replace_no%
 End If
 n_(0) = -5000
 If is_dparal(t_l.line_no(0), t_l.line_no(1), tn%, n_(0), n_(1), n_(2), _
            t_l.line_no(0), t_l.line_no(1)) Then
   If i% = tn% Then
    Dparal(i%).data(0).data0 = t_l
   Else
    If Dparal(i%).data(0).data0.record.data0.condition_data.level < Dparal(tn%).data(0).data0.record.data0.condition_data.level Then
     Call remove_record(paral_, tn%, 0)
    Else
     Call remove_record(paral_, i%, 0)
    End If
   End If
 Else
  If is_line_line_intersect(t_l.line_no(0), t_l.line_no(1), 0, 0, False) > 0 Then
   ' 有公共点
   Call remove_record(paral_, i%, 0)
  Else
  Call search_for_paral(Dparal(i%).data(0).data0, 0, n1_(0), 1)
  Call search_for_paral(Dparal(i%).data(0).data0, 1, n1_(1), 1)
'  Call search_for_paral(Dparal(i%).data(0), 2, n1_(2), 1)
  Dparal(i%).data(0).data0.line_no(0) = t_l.line_no(0)
  Dparal(i%).data(0).data0.line_no(1) = t_l.line_no(1)
  For j% = 0 To 1
   If n1_(j%) < n_(j%) Then
    For l% = n1_(j%) + 1 To n_(j%) - 1
     Dparal(l%).data(0).data0.record.data1.index.i(j%) = Dparal(l% + 1).data(0).data0.record.data1.index.i(j%)
    Next l%
     Dparal(n_(j%)).data(0).data0.record.data1.index.i(j%) = i%
   ElseIf n1_(j%) > n_(j%) Then
    For l% = n1_(j%) + 1 To n_(j%) + 2 Step -1
     Dparal(l%).data(0).data0.record.data1.index.i(j%) = Dparal(l% - 1).data(0).data0.record.data1.index.i(j%)
    Next l%
     Dparal(n_(j%) + 1).data(0).data0.record.data1.index.i(j%) = i%
   End If
  Next j%
  End If
  End If
 End If
 End If
Next k%
For k% = 0 To last_conclusion - 1
 If conclusion_data(k%).ty = tangent_line_ And conclusion_data(k%).no(0) = 0 Then
    If con_tangent_line(k%).data(0).line_no = no1% Or con_tangent_line(k%).data(0).line_no = no2% Then
        con_tangent_line(k%).data(0).line_no = replace_no%
    End If
 End If
Next k%

For k% = 1 + last_conditions.last_cond(0).tangent_line_no To last_conditions.last_cond(1).tangent_line_no
 i% = tangent_line(k%).data(0).record.data1.index.i(0)
 If tangent_line(i%).record_.no_reduce < 255 Then
 If tangent_line(i%).data(0).line_no = no1% Or tangent_line(i%).data(0).line_no = no2% Then
  record_0.data0.condition_data.condition_no = 0 ' record0
  If is_tangent_line(replace_no%, tangent_line(i%).data(0).poi(0), _
           tangent_line(i%).data(0).ele(0), tangent_line(i%).data(0).poi(1), _
           tangent_line(i%).data(0).ele(1), tan_l, tn%, 0, 0, record_0) Then
            If i% = tn% Then
             tangent_line(i%).data(0) = tan_l
            Else
             If tangent_line(i%).data(0).record.data0.condition_data.level < tangent_line(tn%).data(0).record.data0.condition_data.level Then
              Call remove_record(tangent_line_, tn%, 0)
             Else
              Call remove_record(tangent_line_, i%, 0)
             End If
            End If
  Else
           Call remove_record(tangent_line_, i%, 0)
             temp_record.record_data = re
             Call add_conditions_to_record(tangent_line_, i%, 0, 0, temp_record.record_data.data0.condition_data)
          Call set_tangent_line(replace_no%, tangent_line(i%).data(0).poi(0), _
           tangent_line(i%).data(0).ele(0).no, tangent_line(i%).data(0).poi(1), _
            tangent_line(i%).data(0).ele(1).no, temp_record, 0, 0)
  End If
 End If
End If
Next k%
For k% = 0 To last_conclusion - 1
 If conclusion_data(k%).ty = verti_ And conclusion_data(k%).no(0) = 0 Then
    If con_verti(k%).data(0).line_no(0) = no1% Or con_verti(k%).data(0).line_no(0) = no2% Then
        con_verti(k%).data(0).line_no(0) = replace_no%
    ElseIf con_verti(k%).data(0).line_no(1) = no1% Or con_verti(k%).data(0).line_no(1) = no2% Then
        con_verti(k%).data(0).line_no(1) = replace_no%
    End If
    If con_verti(k%).data(0).line_no(0) > con_verti(k%).data(0).line_no(1) Then
     Call exchange_two_integer(con_verti(k%).data(0).line_no(0), con_verti(k%).data(0).line_no(1))
    End If
 End If
Next k%
For k% = 1 + last_conditions.last_cond(0).verti_no To last_conditions.last_cond(1).verti_no
i% = Dverti(k%).data(0).record.data1.index.i(0)
 If Dverti(i%).record_.no_reduce < 255 Then
 t_l = Dverti(i%).data(0)
 If t_l.line_no(0) = no1% Or t_l.line_no(0) = no2% Or _
       t_l.line_no(1) = no1% Or t_l.line_no(1) = no2% Then
  If t_l.line_no(0) = no1% Or t_l.line_no(0) = no2% Then
     t_l.line_no(0) = replace_no%
  End If
  If t_l.line_no(1) = no1% Or t_l.line_no(1) = no2% Then
    t_l.line_no(1) = replace_no%
  End If '
  
  n_(0) = -5000
 If is_dverti(t_l.line_no(0), t_l.line_no(1), tn%, n_(0), n_(1), n_(2), t_l.line_no(0), t_l.line_no(1)) Then
   If i% = tn% Then
    Dverti(i%).data(0) = t_l
   Else
    If Dverti(i%).data(0).record.data0.condition_data.level < Dverti(tn%).data(0).record.data0.condition_data.level Then
     Call remove_record(verti_, tn%, 0)
    Else
     Call remove_record(verti_, i%, 0)
    End If
   End If
 Else
  Call search_for_verti(Dverti(i%).data(0), 0, n1_(0), 1)
  Call search_for_verti(Dverti(i%).data(0), 1, n1_(1), 1)
  Dverti(i%).data(0).line_no(0) = t_l.line_no(0)
  Dverti(i%).data(0).line_no(1) = t_l.line_no(1)
  For j% = 0 To 1
   If n1_(j%) < n_(j%) Then
    For l% = n1_(j%) + 1 To n_(j%) - 1
     Dverti(l%).data(0).record.data1.index.i(j%) = Dverti(l% + 1).data(0).record.data1.index.i(j%)
    Next l%
     Dverti(n_(j%)).data(0).record.data1.index.i(j%) = i%
   ElseIf n1_(j%) > n_(j%) Then
    For l% = n1_(j%) + 1 To n_(j%) + 2 Step -1
     Dverti(l%).data(0).record.data1.index.i(j%) = Dverti(l% - 1).data(0).record.data1.index.i(j%)
    Next l%
     Dverti(n_(j%) + 1).data(0).record.data1.index.i(j%) = i%
   End If
  Next j%

    Call remove_record(verti_, i%, 0)
     temp_record.record_data = re
      Call add_conditions_to_record(verti_, i%, 0, 0, temp_record.record_data.data0.condition_data)
       Call set_dverti(t_l.line_no(0), t_l.line_no(1), temp_record, 0, 0, False)
 End If
 End If
 End If
 Next k%
'***************************
For k% = 0 To last_conclusion - 1
 If conclusion_data(k%).ty = dpoint_pair_ And conclusion_data(k%).no(0) = 0 Then
    If con_dpoint_pair(k%).data(0).line_no(0) = no1% Or con_dpoint_pair(k%).data(0).line_no(0) = no2% Or _
        con_dpoint_pair(k%).data(0).line_no(0) = replace_no% Or con_dpoint_pair(k%).data(0).line_no(1) = no1% Or _
         con_dpoint_pair(k%).data(0).line_no(1) = no2% Or con_dpoint_pair(k%).data(0).line_no(0) = replace_no% Or _
       con_dpoint_pair(k%).data(0).line_no(2) = no1% Or con_dpoint_pair(k%).data(0).line_no(2) = no2% Or _
        con_dpoint_pair(k%).data(0).line_no(2) = replace_no% Or con_dpoint_pair(k%).data(0).line_no(3) = no1% Or _
          con_dpoint_pair(k%).data(0).line_no(3) = no2% Or con_dpoint_pair(k%).data(0).line_no(3) = replace_no% Then
    Call is_point_pair(con_dpoint_pair(k%).data(0).poi(0), con_dpoint_pair(k%).data(0).poi(1), _
           con_dpoint_pair(k%).data(0).poi(2), con_dpoint_pair(k%).data(0).poi(3), _
            con_dpoint_pair(k%).data(0).poi(4), con_dpoint_pair(k%).data(0).poi(5), _
             con_dpoint_pair(k%).data(0).poi(6), con_dpoint_pair(k%).data(0).poi(7), _
              0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                con_dpoint_pair(k%).data(0), 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "", temp_record.record_data)
      End If
  End If
Next k%
For k% = 1 + last_conditions.last_cond(0).dpoint_pair_no To last_conditions.last_cond(1).dpoint_pair_no
 i% = Ddpoint_pair(k%).data(0).record.data1.index.i(0)
ty = False
For j% = 0 To 3
If Ddpoint_pair(i%).data(0).data0.line_no(j%) = no1% Or _
    Ddpoint_pair(i%).data(0).data0.line_no(j%) = no2% Or _
     Ddpoint_pair(i%).data(0).data0.line_no(j%) = replace_no% Then
     '含合并的直线
 ty = max_b(ty, True)
End If
Next j%
If ty Then
tn% = 0
 record_0.data0.condition_data.condition_no = 0 ' record0
  n_(0) = -5000
If is_point_pair(Ddpoint_pair(i%).data(0).data0.poi(0), _
   Ddpoint_pair(i%).data(0).data0.poi(1), _
    Ddpoint_pair(i%).data(0).data0.poi(2), _
     Ddpoint_pair(i%).data(0).data0.poi(3), _
      Ddpoint_pair(i%).data(0).data0.poi(4), _
       Ddpoint_pair(i%).data(0).data0.poi(5), _
        Ddpoint_pair(i%).data(0).data0.poi(6), _
         Ddpoint_pair(i%).data(0).data0.poi(7), _
   0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, tn%, n_(0), _
    n_(1), n_(2), n_(3), n_(4), n_(5), _
     dp_data0, 0, 0, 0, 0, 0, 0, 0, 0, _
      0, "", "", record_0) Then
If tn% <> i% Then
 If Ddpoint_pair(i%).data(0).record.data0.condition_data.level < Ddpoint_pair(tn%).data(0).record.data0.condition_data.level Then
  Call remove_record(dpoint_pair_, tn%, 0)
 Else
  Call remove_record(dpoint_pair_, i%, 0)
 End If
Else
    Ddpoint_pair(i%).data(0).data0 = dp_data0
End If
Else
Call search_for_point_pair(Ddpoint_pair(i%).data(0).data0, 0, n1_(0), 1)
Call search_for_point_pair(Ddpoint_pair(i%).data(0).data0, 1, n1_(1), 1)
Call search_for_point_pair(Ddpoint_pair(i%).data(0).data0, 2, n1_(2), 1)
Call search_for_point_pair(Ddpoint_pair(i%).data(0).data0, 3, n1_(3), 1)
Call search_for_point_pair(Ddpoint_pair(i%).data(0).data0, 4, n1_(4), 1)
Call search_for_point_pair(Ddpoint_pair(i%).data(0).data0, 5, n1_(5), 1)
'Call search_for_point_pair(Ddpoint_pair(i%).data(0).data0, 6, n1_(6), 1)
'Call search_for_point_pair(Ddpoint_pair(i%).data(0).data0, 7, n1_(7), 1)
Ddpoint_pair(i%).data(0).data0 = dp_data0
For j% = 0 To 5
 If n1_(j%) < n_(j%) Then
  For l% = n1_(j%) + 1 To n_(j%) - 1
   Ddpoint_pair(l%).data(0).record.data1.index.i(j%) = Ddpoint_pair(l% + 1).data(0).record.data1.index.i(j%)
  Next l%
   Ddpoint_pair(n_(j%)).data(0).record.data1.index.i(j%) = i%
 ElseIf n1_(j%) > n_(j%) Then
  For l% = n1_(j%) + 1 To n_(j%) + 2 Step -1
   Ddpoint_pair(l%).data(0).record.data1.index.i(j%) = Ddpoint_pair(l% - 1).data(0).record.data1.index.i(j%)
  Next l%
   Ddpoint_pair(n_(j%) + 1).data(0).record.data1.index.i(j%) = i%
 End If
Next j%
End If
End If
Next k%
'******************************
For k% = 0 To last_conclusion - 1
 If conclusion_data(k%).ty = relation_ And conclusion_data(k%).no(0) = 0 Then
    If con_relation(k%).data(0).line_no(0) = no1% Or con_relation(k%).data(0).line_no(0) = no2% Or _
        con_relation(k%).data(0).line_no(0) = replace_no% Or con_relation(k%).data(0).line_no(1) = no1% Or _
          con_relation(k%).data(0).line_no(1) = no2% Or con_relation(k%).data(0).line_no(1) = replace_no% Then
    Call is_relation(con_relation(k%).data(0).poi(0), con_relation(k%).data(0).poi(1), _
                      con_relation(k%).data(0).poi(2), con_relation(k%).data(0).poi(3), _
                        0, 0, 0, 0, 0, 0, con_relation(k%).data(0).value, 0, 0, 0, 0, 0, con_relation(k%).data(0), 0, 0, 0, _
                          temp_record.record_data.data0.condition_data, 0)
    End If
 End If
Next k%
For k% = 1 + last_conditions.last_cond(0).relation_no To last_conditions.last_cond(1).relation_no
i% = Drelation(k%).data(0).record.data1.index.i(0)
If Drelation(i%).record_.no_reduce < 255 Then
If Drelation(i%).data(0).data0.line_no(0) = no1% Or Drelation(i%).data(0).data0.line_no(1) = no1% Or _
     Drelation(i%).data(0).data0.line_no(0) = no2% Or Drelation(i%).data(0).data0.line_no(1) = no2% Or _
      Drelation(i%).data(0).data0.line_no(0) = replace_no% Or Drelation(i%).data(0).data0.line_no(1) = replace_no% Then
  record_0.data0.condition_data.condition_no = 0 'record0
   n_(0) = -5000
ty = is_relation(Drelation(i%).data(0).data0.poi(0), Drelation(i%).data(0).data0.poi(1), _
       Drelation(i%).data(0).data0.poi(2), Drelation(i%).data(0).data0.poi(3), _
        0, 0, 0, 0, 0, 0, Drelation(i%).data(0).data0.value, tn%, n_(0), _
         n_(1), n_(2), n_(3), relation_data0, 0, 0, 0, record_0.data0.condition_data, 0)
If relation_data0.value = "1" Then
 temp_record.record_data = re
 Call add_conditions_to_record(relation_, i%, 0, 0, temp_record.record_data.data0.condition_data)
 Call set_equal_dline(relation_data0.poi(0), relation_data0.poi(1), _
         relation_data0.poi(2), relation_data0.poi(3), _
          relation_data0.n(0), relation_data0.n(1), _
           relation_data0.n(2), relation_data0.n(3), _
            relation_data0.line_no(0), relation_data0.line_no(1), _
             0, temp_record, 0, 0, 0, 0, 0, False)
          Call remove_record(relation_, i%, 0)
Else
If ty Then
 If tn% <> i% Then
  If Drelation(i%).data(0).record.data0.condition_data.level < Drelation(tn%).data(0).record.data0.condition_data.level Then
   is_remove = remove_record(relation_, tn%, 1)
  Else
   is_remove = remove_record(relation_, i%, 1)
    i% = tn%
  End If
 Else
  Drelation(i%).data(0).data0 = relation_data0
   GoTo simple_dbase_for_line_relation_next
 End If
   'is_remove = remove_record(relation_, i%, 1)
   Drelation(i%).data(0).data0 = relation_data0
   Call add_record(relation_, i%, is_remove)
   simple_dbase_for_line = combine_relation_with_others_condition(i%, 0)
   If simple_dbase_for_line > 1 Then
      Exit Function
   End If
Else
 Call search_for_relation(Drelation(i%).data(0).data0, 0, n1_(0), 1)
 Call search_for_relation(Drelation(i%).data(0).data0, 1, n1_(1), 1)
 Call search_for_relation(Drelation(i%).data(0).data0, 2, n1_(2), 1)
 Call search_for_relation(Drelation(i%).data(0).data0, 3, n1_(3), 1)
  Drelation(i%).data(0).data0 = relation_data0
 For j% = 0 To 3
  If n1_(j%) < n_(j%) Then
   For l% = n1_(j%) + 1 To n_(j%) - 1
    Drelation(l%).data(0).record.data1.index.i(j%) = Drelation(l% + 1).data(0).record.data1.index.i(j%)
   Next l%
    Drelation(n_(j%)).data(0).record.data1.index.i(j%) = i%
  ElseIf n1_(j%) > n_(j%) Then
   For l% = n1_(j%) + 1 To n_(j%) + 2 Step -1
    Drelation(l%).data(0).record.data1.index.i(j%) = Drelation(l% - 1).data(0).record.data1.index.i(j%)
   Next l%
    Drelation(n_(j%) + 1).data(0).record.data1.index.i(j%) = i%
  End If
 Next j%
End If
End If
End If
End If
simple_dbase_for_line_relation_next:
Next k%
For k% = 1 + last_conditions.last_cond(0).item0_no To last_conditions.last_cond(1).item0_no
 i% = item0(k%).data(0).index(0)
If item0(i%).data(0).line_no(0) = no1% Or item0(i%).data(0).line_no(0) = no1% Or _
     item0(i%).data(0).line_no(0) = no2% Or item0(i%).data(0).line_no(1) = no2% Or _
      item0(i%).data(0).line_no(0) = replace_no% Or item0(i%).data(0).line_no(1) = replace_no% Then
  record_0.data0.condition_data.condition_no = 0 ' record0
   n_(0) = 0
If is_item0(item0(i%).data(0).poi(0), item0(i%).data(0).poi(1), _
       item0(i%).data(0).poi(2), item0(i%).data(0).poi(3), _
        item0(i%).data(0).sig, 0, 0, 0, 0, 0, 0, tn%, n_(0), _
         n_(1), n_(2), "", item_0) Then
item0(tn%).data(0).line_no(0) = item_0.line_no(0)
item0(tn%).data(0).line_no(1) = item_0.line_no(1)
item0(tn%).data(0).line_no(2) = item_0.line_no(2)
item0(tn%).data(0).n(0) = item_0.n(0)
item0(tn%).data(0).n(1) = item_0.n(1)
item0(tn%).data(0).n(2) = item_0.n(2)
item0(tn%).data(0).n(3) = item_0.n(3)
item0(tn%).data(0).n(4) = item_0.n(4)
item0(tn%).data(0).n(5) = item_0.n(5)
If tn% <> i% Then
   is_remove = remove_record(item0_, i%, 1)
   item0(i%).data(0) = item0(tn%).data(0)
    Call add_record(item0_, i%, is_remove)
 End If
Else
Call search_for_item0(item0(i%).data(0), 0, n1_(0), 1)
Call search_for_item0(item0(i%).data(0), 1, n1_(1), 1)
Call search_for_item0(item0(i%).data(0), 2, n1_(2), 1)
'Call search_for_item0(item0(i%).data(0), 3, n1_(3), 1)
item0(i%).data(0).line_no(0) = item_0.line_no(0)
item0(i%).data(0).line_no(1) = item_0.line_no(1)
item0(i%).data(0).line_no(2) = item_0.line_no(2)
item0(i%).data(0).poi(0) = item_0.poi(0)
item0(i%).data(0).poi(1) = item_0.poi(1)
item0(i%).data(0).poi(2) = item_0.poi(2)
item0(i%).data(0).poi(3) = item_0.poi(3)
item0(i%).data(0).poi(4) = item_0.poi(4)
item0(i%).data(0).poi(5) = item_0.poi(5)
item0(i%).data(0).n(0) = item_0.n(0)
item0(i%).data(0).n(1) = item_0.n(1)
item0(i%).data(0).n(2) = item_0.n(2)
item0(i%).data(0).n(3) = item_0.n(3)
item0(i%).data(0).n(4) = item_0.n(4)
item0(i%).data(0).n(5) = item_0.n(5)
For j% = 0 To 2
 If n1_(j%) < n_(j%) Then
  For l% = n_(j%) + 1 To n_(j%) - 1
   item0(l%).data(0).index(j%) = item0(l% + 1).data(0).index(j%)
  Next l%
  item0(n_(j%)).data(0).index(j%) = i%
 ElseIf n1_(j%) > n_(j%) Then
  For l% = n_(j%) + 1 To n_(j%) + 2 Step -1
   item0(l%).data(0).index(j%) = item0(l% - 1).data(0).index(j%)
  Next l%
  item0(n_(j%) + 1).data(0).index(j%) = i%
 End If
Next j%
'Call remove_record(item0_, i%, 0)
'Call set_item0(,(item0_, i%)
End If
End If
Next k%
simple_dbase_for_line:
End Function
Private Sub remove_line_form_data(ByVal l%)
'　消去l%并改变线号
Dim i%, j%, k%
Dim tl%
tl% = l%
i% = last_conditions.last_cond(1).line_from_two_point_no
Do While i% > 0
If Dtwo_point_line(i%).data(0).line_no = tl% Then
  Dtwo_point_line(i%).data(0).line_no = 0
End If
Do While Dtwo_point_line(last_conditions.last_cond(1).line_from_two_point_no).data(0).line_no = 0 And _
     last_conditions.last_cond(1).line_from_two_point_no > 0
      last_conditions.last_cond(1).line_from_two_point_no = last_conditions.last_cond(1).line_from_two_point_no - 1
Loop
i% = min(i% - 1, last_conditions.last_cond(1).line_from_two_point_no)
Loop
For i% = 1 To last_conditions.last_cond(1).line_no '12.30
'***********************************************
 If i% <> tl% Then
'******************************************
'消除平行线中的l%
'*******************************************
 For j% = 1 To m_lin(i%).data(0).in_paral(0).line_no
  If tl% = m_lin(i%).data(0).in_paral(j%).line_no Then
   For k% = j% To m_lin(i%).data(0).in_paral(0).line_no - 1
   m_lin(i%).data(0).in_paral(k%) = m_lin(i%).data(0).in_paral(k% - 1)
   'lin(i%).data(0)..paral_no(k%) = lin(i%).data(0).paral_no(k% - 1)
   Next k%
   m_lin(i%).data(0).in_paral(0).line_no = m_lin(i%).data(0).in_paral(0).line_no - 1
     GoTo remove_line_mark1
  End If
 Next j%

'***************************************************
remove_line_mark1:
 For j% = 1 To m_lin(i%).data(0).in_verti(0).line_no
 '*******************************************************
  If tl% = m_lin(i%).data(0).in_verti(j%).line_no Then
   For k% = j% To m_lin(i%).data(0).in_verti(0).line_no - 1
    m_lin(i%).data(0).in_verti(k%) = m_lin(i%).data(0).in_verti(k% + 1)
    'lin(i%).data(0).verti_no(k%) = lin(i%).data(0).verti_no(k% + 1)
   Next k%
   m_lin(i%).data(0).in_verti(0).line_no = m_lin(i%).data(0).in_verti(0).line_no - 1
     GoTo remove_line_mark2
 End If
 '******************************************
 Next j%
remove_line_mark2:
 End If
Next i%
m_lin(tl%).data(0).data0.poi(0) = 0
 m_lin(tl%).data(0).data0.poi(1) = 0
For i% = 0 To 10
  m_lin(tl%).data(0).data0.in_point(i%) = 0
Next i%
For i% = 1 To last_conditions.last_cond(1).point_no
   j% = 1
  Do While j% <= m_poi(i%).data(0).in_line(0)
   If m_poi(i%).data(0).in_line(j%) = l% Then
       m_poi(i%).data(0).in_line(0) = m_poi(i%).data(0).in_line(0) - 1
       For k% = j% To m_poi(i%).data(0).in_line(0)
          m_poi(i%).data(0).in_line(k%) = m_poi(i%).data(0).in_line(k% + 1)
           If m_poi(i%).data(0).in_line(k%) > l% Then
             m_poi(i%).data(0).in_line(k%) = m_poi(i%).data(0).in_line(k%) - 1
           End If
       Next k%
       GoTo remove_line_mark3
   ElseIf m_poi(i%).data(0).in_line(j%) > l% Then
       m_poi(i%).data(0).in_line(j%) = m_poi(i%).data(0).in_line(j%) - 1
    j% = j% + 1
   Else
    j% = j% + 1
   End If
  Loop
remove_line_mark3:
Next i%
' last_conditions.last_cond(1).line_no = last_conditions.last_cond(1).line_no - 1 '***
 ' For i% = tl% To last_conditions.last_cond(1).line_no '***
  ' Lin(i%) = Lin(i% + 1)
   ' Next i%
    'Lin(last_conditions.last_cond(1).line_no + 1).data(0).data0.in_point(0) = 0
  '********************************************
 'For i% = 0 To 3
 'If temp_line(i%) > tl% Then
  'temp_line(i%) = temp_line(i%) - 1
 'End If
 'Next i%
' Call init_line0(last_conditions.last_cond(1).line_no + 1)
 For i% = 1 To last_conditions.last_cond(1).angle_no
  If angle(i%).data(0).line_no(0) = tl% Or angle(i%).data(0).line_no(1) = tl% Then
   angle(i%).data(0).line_no(0) = 0
    angle(i%).data(0).line_no(1) = 0
  End If
 Next i%
 For i% = 1 To last_conditions.last_cond(1).paral_no
  If Dparal(i%).data(0).data0.line_no(0) = tl% Or Dparal(i%).data(0).data0.line_no(1) = tl% Then
   Dparal(i%).data(0).data0.line_no(0) = 0
    Dparal(i%).data(0).data0.line_no(1) = 0
  End If
 Next i%
  For i% = 1 To last_conditions.last_cond(1).verti_no
  If Dverti(i%).data(0).line_no(0) = tl% Or Dverti(i%).data(0).line_no(1) = tl% Then
   Dverti(i%).data(0).line_no(0) = 0
    Dverti(i%).data(0).line_no(1) = 0
  End If
 Next i%

 ' For i% = 1 To last_conditions.last_cond(1).point_no
 '
 '  For j% = 1 To poi(i%).data(0).in_line(0)
 '   If poi(i%).data(0).in_line(j%) = tl% Then
 '    poi(i%).data(0).in_line(0) = poi(i%).data(0).in_line(0) - 1
 '    For k% = j% To poi(i%).data(0).in_line(0)
 '     poi(i%).data(0).in_line(k%) = poi(i%).data(0).in_line(k% + 1)
 '      Next k%
 '       j% = j% - 1
 '    ElseIf poi(i%).data(0).in_line(j%) > tl% Then
 '    poi(i%).data(0).in_line(j%) = poi(i%).data(0).in_line(j%) - 1
 '   End If
 '  Next j%
 ' Next i%
'**********************************************************
 'End If

End Sub

Public Sub delete_line(ByVal l%)
'消除线
 Dim i%, j%, k%
  For i% = 1 To last_conditions.last_cond(1).line_no
   If i% > l% Then
    Call move_line_data(i%, i% - 1) 'l%后的数据,前移
   ElseIf i% = l% Then
     last_conditions.last_cond(1).line_no = last_conditions.last_cond(1).line_no - 1
   End If
  Next i%
'*******************************************************************************************
  Call remove_line_form_data(l%)
  Call delete_line_from_poi(l%)
'For i% = 1 To last_conditions.last_cond(1).point_no
 'k% = 1
 'For j% = 1 To m_poi(i%).data(0).in_line(0)
  'If m_poi(i%).data(0).in_line(j%) < l% Then
   ' poi(i%).data(0).in_line(k%) = poi(i%).data(0).in_line(j%)
  '   k% = k% + 1
 'ElseIf poi(i%).data(0).in_line(j%) > l% Then
  ' poi(i%).data(0).in_line(k%) = poi(i%).data(0).in_line(j%) - 1
   ' k% = k% + 1
 'Else
 '  poi(i%).data(0).in_line(0) = poi(i%).data(0).in_line(0) - 1
 'End If
 'Next j%
'Next i%
'For i% = 0 To 3
 'If temp_line(i%) > l% Then
  'temp_line(i%) = temp_line(i%) - 1
 'End If
 'Next i%
End Sub

Private Sub change_line_(m_p%)    '由点的边化,导致圆变化
Dim i%, l%, c%
 If m_p% > 0 Then
  For i% = 1 To m_poi(m_p%).data(0).in_line(0)
     l% = m_poi(m_p%).data(0).in_line(i%)
'            Call change_point_by_line(l%)
        Call simple_line(l%, m_p%)
  Next i%
    For i% = 1 To m_poi(m_p%).data(0).in_circle(0)
     c% = m_poi(m_p%).data(0).in_circle(i%)
'            Call change_point_by_line(l%)
        Call simple_circle(c%, m_p%)
  Next i%
 End If
End Sub
Private Sub simple_line(ByVal line_no%, ByVal point_no%)  ', t As Boolean) '12.30
Dim l As line_data_type
Dim old_l As line_data_type
Dim i%, j%, tp%
Dim is_change_ As Boolean
'　重排序
l = m_lin(line_no%).data(0)
old_l = l
If (m_poi(l.data0.poi(0)).data(0).data0.coordinate.X = 10000 And _
      m_poi(l.data0.poi(0)).data(0).data0.coordinate.Y = 10000) Or _
       (m_poi(l.data0.poi(1)).data(0).data0.coordinate.X = 10000 And _
         m_poi(l.data0.poi(1)).data(0).data0.coordinate.Y = 10000) Then
Exit Sub
End If
If is_point_in_line3(point_no%, l.data0, i%) Then
If i% = 1 Then
         is_change_ = True
 If compare_two_point(m_poi(point_no%).data(0).data0.coordinate, _
              m_poi(l.data0.in_point(2)).data(0).data0.coordinate, 0, 0, 6) = -1 Then
   Call exchange_two_integer(l.data0.in_point(1), l.data0.in_point(2))
   Call exchange_two_integer(l.data0.in_point(1), l.data0.in_point(2))
   If l.data0.in_point(1) < 0 And l.data0.in_point(2) > 0 Then
      l.data0.in_point(2) = -l.data0.in_point(2)
   ElseIf l.data0.in_point(1) > 0 And l.data0.in_point(2) < 0 Then
      l.data0.in_point(1) = -l.data0.in_point(1)
      l.data0.in_point(2) = -l.data0.in_point(2)
   End If
 End If
ElseIf i% = l.data0.in_point(0) Then
        is_change_ = True
 If compare_two_point(m_poi(l.data0.in_point(l.data0.in_point(0) - 1)).data(0).data0.coordinate, _
                 m_poi(point_no%).data(0).data0.coordinate, 0, 0, 6) = -1 Then
   Call exchange_two_integer(l.data0.in_point(l.data0.in_point(0) - 1), l.data0.in_point(l.data0.in_point(0)))
   Call exchange_two_integer(l.data0.in_point(l.data0.in_point(0) - 1), l.data0.in_point(l.data0.in_point(0)))
   If l.data0.in_point(l.data0.in_point(0) - 1) > 0 And l.data0.in_point(l.data0.in_point(0)) < 0 Then
      l.data0.in_point(l.data0.in_point(0)) = -l.data0.in_point(l.data0.in_point(0))
      l.data0.in_point(l.data0.in_point(0) - 1) = -l.data0.in_point(l.data0.in_point(0) - 1)
   End If
 End If
Else
 If compare_two_point(m_poi(point_no%).data(0).data0.coordinate, _
             m_poi(l.data0.in_point(i% + 1)).data(0).data0.coordinate, 0, 0, 6) = -1 Then
              is_change_ = True
   Call exchange_two_integer(l.data0.in_point(i%), l.data0.in_point(i% + 1))
   Call exchange_two_integer(l.data0.in_point(i%), l.data0.in_point(i% + 1))
    If l.data0.in_point(i%) < 0 And l.data0.in_point(i% + 1) > 0 Then
       l.data0.in_point(i% + 1) = -l.data0.in_point(i% + 1)
           is_change_ = True
    ElseIf l.data0.in_point(i%) > 0 And l.data0.in_point(i% + 1) < 0 Then
       l.data0.in_point(i%) = -l.data0.in_point(i%)
    End If
 ElseIf compare_two_point(m_poi(l.data0.in_point(i% - 1)).data(0).data0.coordinate, _
                  m_poi(point_no%).data(0).data0.coordinate, 0, 0, 6) = -1 Then
           is_change_ = True
   Call exchange_two_integer(l.data0.in_point(i%), l.data0.in_point(i% - 1))
   Call exchange_two_integer(l.data0.in_point(i%), l.data0.in_point(i% - 1))
     If (l.data0.in_point(i% - 1) < 0 And l.data0.in_point(i%) > 0) Or _
          (l.data0.in_point(i% - 1) > 0 And l.data0.in_point(i%) < 0) Then
      l.data0.in_point(i%) = -l.data0.in_point(i%)
       is_change_ = True
     End If
 End If
End If
End If
l.data0.poi(0) = l.data0.in_point(1)
l.data0.poi(1) = l.data0.in_point(l.data0.in_point(0))
If l.data0.visible > 0 Then
If is_change_ Then
 l.is_change = True
Else
 If l.data0.poi(0) = point_no% Or l.data0.poi(1) = point_no% Then
  l.is_change = True
 End If
End If
End If
'****************************************************************
m_lin(line_no%).data(1) = m_lin(line_no%).data(0)
m_lin(line_no%).data(0) = l
If old_l.data0.poi(0) = point_no% Or old_l.data0.poi(1) = point_no% Or _
     l.data0.poi(0) = point_no% Or l.data0.poi(1) = point_no% Or _
       l.data0.poi(0) = point_no% Or l.data0.poi(1) = point_no% Then
m_lin(line_no%).data(0).is_change = 255
Call C_display_picture.re_draw_line(line_no%)
End If
End Sub
Private Sub simple_circle(ByVal circle_no%, point_no%) ', t As Boolean) '12.30
If m_Circ(circle_no%).data(0).circle_type = 1 Then
   If m_Circ(circle_no%).data(0).data0.center = point_no% Then
      m_Circ(circle_no%).data(0).data0.c_coord = _
          m_poi(m_Circ(circle_no).data(0).data0.center).data(0).data0.coordinate
   ElseIf m_Circ(circle_no%).data(0).data0.in_point(1) = point_no% Then
   Else
     Exit Sub
   End If
    m_Circ(circle_no).data(0).data0.radii = _
          sqr((m_poi(m_Circ(circle_no).data(0).data0.in_point(1)).data(0).data0.coordinate.X - _
               m_poi(m_Circ(circle_no).data(0).data0.center).data(0).data0.coordinate.X) ^ 2 + _
              (m_poi(m_Circ(circle_no).data(0).data0.in_point(1)).data(0).data0.coordinate.Y - _
               m_poi(m_Circ(circle_no).data(0).data0.center).data(0).data0.coordinate.Y) ^ 2)
          m_Circ(circle_no%).data(0).is_change = True
ElseIf m_Circ(circle_no%).data(0).circle_type = 2 Then
   If m_Circ(circle_no%).data(0).data0.in_point(1) = point_no% Or _
      m_Circ(circle_no%).data(0).data0.in_point(2) = point_no% Or _
      m_Circ(circle_no%).data(0).data0.in_point(3) = point_no% Then
   m_Circ(circle_no).data(0).data0.radii = _
          circle_radii0(m_poi(m_Circ(circle_no).data(0).data0.in_point(1)).data(0).data0.coordinate, _
                        m_poi(m_Circ(circle_no).data(0).data0.in_point(2)).data(0).data0.coordinate, _
                        m_poi(m_Circ(circle_no).data(0).data0.in_point(3)).data(0).data0.coordinate, _
                        m_Circ(circle_no).data(0).data0.c_coord)
  Call set_point_coordinate(m_Circ(circle_no).data(0).data0.center, _
         m_Circ(circle_no).data(0).data0.c_coord, True)
   End If
End If
End Sub

Public Sub set_temp_line(ByVal n%, ByVal l%, p As POINTAPI)
Dim tl As line_data_type
Dim dr(1) As Integer
If l% <> 0 Then
tl = m_lin(l%).data(0)
 dr(0) = compare_two_point(p, m_poi(tl.data0.poi(0)).data(0).data0.coordinate, 0, 0, 0)
 dr(1) = compare_two_point(p, m_poi(tl.data0.poi(1)).data(0).data0.coordinate, 0, 0, 0)
If tl.data0.visible > 0 Then '可见线段
If dr(0) = 1 And (dr(1) = 1 Or dr(1) = 0) Then '在线段前
    If m_poi(tl.data0.poi(0)).data(0).data0.visible = 0 Then
       If m_poi(tl.data0.poi(1)).data(0).data0.visible = 1 Then
        tangent_line(n%).p(0) = p
        tangent_line(n%).p(1) = m_poi(tl.data0.poi(1)).data(0).data0.coordinate
       Else
        tangent_line(n%).p(0) = p
        tangent_line(n%).p(1) = m_poi(tl.in_point( _
                         tl.in_point(0) - 1)).data(0).data0.coordinate
       End If
    Else
        tangent_line(n%).p(0) = p
        tangent_line(n%).p(1) = m_poi(tl.data0.poi(0)).data(0).data0.coordinate
    End If
      tangent_line(n%).visible = 1
ElseIf (dr(0) = -1 Or dr(0) = 0) And (dr(1) = 1 Or dr(1) = 0) Then
    If m_poi(tl.data0.poi(0)).data(0).data0.visible = 0 Then
       If m_poi(tl.data0.poi(1)).data(0).data0.visible > 0 Then
        tangent_line(n%).p(0) = p
        tangent_line(n%).p(1) = m_poi(tl.data0.poi(1)).data(0).data0.coordinate
       Else
        If tl.in_point(0) > 2 Then
        tangent_line(n%).p(0) = p
        tangent_line(n%).p(1) = m_poi(tl.in_point( _
                         tl.in_point(0) - 1)).data(0).data0.coordinate
        Else
         tangent_line(n%).p(0) = m_poi(tl.data0.poi(0)).data(0).data0.coordinate
         tangent_line(n%).p(1) = m_poi(tl.data0.poi(1)).data(0).data0.coordinate
        End If
       End If
        tangent_line(n%).visible = 1
    ElseIf m_poi(tl.data0.poi(0)).data(0).data0.visible > 0 Then
      If m_poi(tl.data0.poi(1)).data(0).data0.visible = 0 Then
       tangent_line(n%).p(0) = m_poi(tl.data0.poi(0)).data(0).data0.coordinate
       tangent_line(n%).p(1) = p
      Else
       Exit Sub
      End If
       tangent_line(n%).visible = 1
    ElseIf m_lin(l%).data(0).data0.total_color = 15 Then '显示,透明
      tangent_line(n%).p(0) = m_poi(tl.data0.poi(0)).data(0).data0.coordinate
      tangent_line(n%).p(1) = m_poi(tl.data0.poi(1)).data(0).data0.coordinate
      tangent_line(n%).visible = 1
    End If
ElseIf dr(0) = -1 And (dr(1) = -1 Or dr(0) = 0) Then
    If m_poi(tl.data0.poi(1)).data(0).data0.visible = 0 Then
      If m_poi(tl.data0.poi(0)).data(0).data0.visible > 0 Then
      tangent_line(n%).p(0) = m_poi(tl.data0.poi(0)).data(0).data0.coordinate
      tangent_line(n%).p(1) = p
      Else
      tangent_line(n%).p(0) = m_poi(tl.in_point(2)).data(0).data0.coordinate
      tangent_line(n%).p(1) = p
      End If
    Else
       tangent_line(n%).p(0) = m_poi(tl.data0.poi(1)).data(0).data0.coordinate
       tangent_line(n%).p(1) = p
    End If
      tangent_line(n%).visible = 1
End If
 Else
  If (dr(0) = 1 Or dr(0) = 0) And dr(1) = 1 Then
    tangent_line(n%).p(0) = m_poi(tl.data0.poi(1)).data(0).data0.coordinate
    tangent_line(n%).p(1) = p
    tangent_line(n%).visible = 1
  ElseIf (dr(0) = -1 Or dr(0) = 0) And (dr(1) = -1 Or dr(1) = 0) Then
  tangent_line(n%).p(0) = p
   tangent_line(n%).p(1) = m_poi(tl.data0.poi(0)).data(0).data0.coordinate
   tangent_line(n%).visible = 1
  ElseIf dr(0) = -1 And (dr(1) = -1 Or dr(1) = 0) Then
  tangent_line(n%).p(0) = m_poi(tl.data0.poi(0)).data(0).data0.coordinate
    tangent_line(n%).p(1) = m_poi(tl.data0.poi(1)).data(0).data0.coordinate
     tangent_line(n%).visible = 1
  End If
 End If
 End If
End Sub
Public Sub set_circle_color(circle_no%, color As Byte)
 Call C_display_picture.set_m_circle_color(circle_no%, color)
End Sub
Public Sub set_aid_point(p%, coord As POINTAPI, vi As Byte)
Dim i%
If p% = 0 Then
      last_aid_point = last_aid_point + 1
       p% = 100 - last_aid_point
End If
For i% = 0 To 15
If aid_point(i%) = p% Then
 Exit Sub
End If
Next i%
For i% = 0 To 7
If aid_point(i%) = 0 Then
aid_point(i%) = p%
GoTo set_aid_point_mark0
End If
Next i%
set_aid_point_mark0:
Call m_point_number(coord, aid_condition, 1, fill_color, "", condition_type0, condition_type0, 0, True)
End Sub


Public Function change_picture_(mp%, ty As Byte) As Boolean
Dim i% 'ty=1 显示轨迹
Dim change_p As Boolean
'yidian_no = mp%
last_change_point = 0
   For i% = 1 To C_display_wenti.m_display_string.Count
    change_picture_ = change_picture(i%, mp%)
   Next i%
   change_picture_ = True
   Call set_new_picture
   If last_conditions.last_cond(1).last_view_point_no > 0 Then
    Draw_form.DrawMode = 13
    For i% = 1 To last_conditions.last_cond(1).last_view_point_no
    Draw_form.Line (m_poi(view_point(i%).poi).data(0).data0.coordinate.X, _
     m_poi(view_point(i%).poi).data(0).data0.coordinate.Y)- _
       (view_point(i%).old_coordinate.X, view_point(i%).old_coordinate.Y), QBColor(12)
        view_point(i%).old_coordinate = m_poi(view_point(i%).poi).data(0).data0.coordinate
    Next i%
   End If
       Draw_form.DrawMode = 10
End Function
Public Sub change_picture_0(ByVal S_p%, ByVal E_p%, num%)
Dim i% 'ty=1 显示轨迹
If S_p% > 0 And E_p% > 0 And S_p% < E_p% Then
last_conditions.last_cond(1).change_picture_step = _
 last_conditions.last_cond(1).change_picture_step + 1
For i% = 1 To last_conditions.last_cond(1).point_no
    m_poi(i%).data(last_conditions.last_cond(1).change_picture_step) = _
      m_poi(i%).data(0)
    m_poi(i%).data(0).is_change = False
Next i%
    m_poi(S_p%).data(0).is_change = True
For i% = 1 To last_conditions.last_cond(1).line_no
    m_lin(i%).data(last_conditions.last_cond(1).change_picture_step) = _
      m_lin(i%).data(0)
    If m_lin(i%).data(0).data0.poi(0) <> S_p% And _
        m_lin(i%).data(0).data0.poi(1) <> S_p% Then
         m_lin(i%).data(0).is_change = 0
    Else
         m_lin(i%).data(0).is_change = 255
    End If
Next i%
'***********************************
   For i% = 1 To num% - 1
     If C_display_wenti.m_point_no(i%, 49) > S_p% Then
         Call change_picture(i%, S_p%)
     End If
   Next i%
'******************************************
For i% = 1 To last_conditions.last_cond(1).point_no
    m_poi(i%).data(0).is_change = _
      m_poi(i%).data(last_conditions.last_cond(1).change_picture_step).is_change Or _
        m_poi(i%).data(0).is_change
Next i%
For i% = 1 To last_conditions.last_cond(1).line_no
    m_lin(i%).data(0).is_change = _
      m_lin(i%).data(last_conditions.last_cond(1).change_picture_step).is_change Or _
        m_lin(i%).data(0).is_change
Next i%
last_conditions.last_cond(1).change_picture_step = _
 last_conditions.last_cond(1).change_picture_step - 1
End If
End Sub

Public Sub change_picture_from_move(ByVal move_point%, new_coordinate As POINTAPI)
Dim i%
Exit Sub
If move_point% > 0 Then
  If regularity_coordinate(move_point%, new_coordinate, new_coordinate) Then
      For i% = 1 To last_conditions.last_cond(1).point_no
       m_poi(i%).data(0).is_change = False
      Next i%
      m_poi(move_point%).data(0).is_change = True
      For i% = 1 To last_conditions.last_cond(1).line_no
         If m_lin(i%).data(0).data0.poi(0) <> move_point% And _
             m_lin(i%).data(0).data0.poi(1) <> move_point% Then
              m_lin(i%).data(0).is_change = 255
         Else
              m_lin(i%).data(0).is_change = 255
         End If
      Next i%
'***********************************
    Call set_point_coordinate(move_point%, new_coordinate, True)
    If change_picture_(move_point%, 0) Then
     Call draw_again0(Draw_form, 1)
    Else
    '图形变化失败,恢复原图
      For i% = 1 To last_conditions.last_cond(1).point_no
       If m_poi(i%).data(0).is_change Then
         m_poi(i%).data(0) = m_poi(i%).data(1)
       End If
      Next i%
      For i% = 1 To last_conditions.last_cond(1).line_no
       If m_lin(i%).data(0).is_change = 255 Then
         m_lin(i%).data(0) = m_lin(i%).data(1)
       End If
      Next i%
   End If
  End If
Else
 For i% = 1 To last_conditions.last_cond(1).point_no
     m_poi(i%).data(0).data0.coordinate.X = m_poi(i%).data(0).data0.coordinate.X + new_coordinate.X
     m_poi(i%).data(0).data0.coordinate.Y = m_poi(i%).data(0).data0.coordinate.Y + new_coordinate.Y
     m_poi(i%).data(0).is_change = True
 Next i%
 For i% = 1 To last_conditions.last_cond(1).line_no
     m_lin(i%).data(0).is_change = 255
 Next i%
' For i% = 1 To last_conditions.last_cond(1).circle_no
'     m_Circ(i%).data(0).data0.c_coord.X = m_Circ(i%).data(0).data0.c_coord.X + new_coordinate.X
'     m_Circ(i%).data(0).data0.c_coord.Y = m_Circ(i%).data(0).data0.c_coord.Y + new_coordinate.Y
'     m_Circ(i%).data(0).is_change = True
' Next i%
     Call draw_again0(Draw_form, 1)
End If
End Sub
Public Sub m_BPset(ob As Object, pcoord As POINTAPI, ByVal n$, _
                      color As Byte) '12.30
Dim imag_position As Integer
On Error GoTo BPset_error
If color = 9 Or color = condition_color Then 'condition_colour=3
   imag_position = 1
ElseIf color = 7 Then
   imag_position = 2
ElseIf color = 12 Or color = conclusion_color Then 'conclusion_color
   imag_position = 3
ElseIf color = 10 Then
   imag_position = 4
Else
   imag_position = 1
End If
If Abs(pcoord.X) < 1000 And Abs(pcoord.Y) < 1000 Then
 If n$ >= "A" And n$ <= "Z" Then
  If imag_position = 1 Then
   If line_width < 3 Then
    ob.PaintPicture Draw_form.ImageList1.ListImages(Asc(n$) - 60).Picture, pcoord.X - 2, pcoord.Y - 2, 16, 18, 0, 0, 16, 18, &H990066
   Else
    ob.PaintPicture Draw_form.ImageList3.ListImages(Asc(n$) - 60).Picture, pcoord.X - 5, pcoord.Y - 5, 32, 32, 0, 0, 32, 32, &H990066
   End If
  Else
  If line_width < 3 Then
   ob.PaintPicture Draw_form.ImageList1.ListImages(imag_position).Picture, pcoord.X - 2, pcoord.Y - 2, 16, 18, 0, 0, 16, 18, &H990066
    Call put_name(ob, pcoord, n$)
  Else
   ob.PaintPicture Draw_form.ImageList3.ListImages(imag_position).Picture, pcoord.X - 5, pcoord.Y - 5, 32, 32, 0, 0, 32, 32, &H990066
    Call put_name(ob, pcoord, n$)
  End If
  End If
 Else
  If line_width < 3 Then
   ob.PaintPicture Draw_form.ImageList1.ListImages(imag_position).Picture, pcoord.X - 2, pcoord.Y - 2, 16, 18, 0, 0, 16, 18, &H990066
  Else
   ob.PaintPicture Draw_form.ImageList3.ListImages(imag_position).Picture, pcoord.X - 5, pcoord.Y - 5, 32, 32, 0, 0, 32, 32, &H990066
  End If
End If
End If
BPset_error:
End Sub
Private Sub put_name(ob As Object, pcoord As POINTAPI, name As String)
If name >= "A" And name <= "Z" Then
 If line_width < 3 Then
 ob.PaintPicture Draw_form.ImageList1.ListImages(1).Picture, pcoord.X - 2, pcoord.Y - 2, 16, 18, 0, 0, 16, 18, &H990066
 ob.PaintPicture Draw_form.ImageList1.ListImages(Asc(name) - 60).Picture, pcoord.X - 2, pcoord.Y - 2, 16, 18, 0, 0, 16, 18, &H990066
Else
 ob.PaintPicture Draw_form.ImageList3.ListImages(1).Picture, pcoord.X - 5, pcoord.Y - 5, 32, 32, 0, 0, 32, 32, &H990066
 ob.PaintPicture Draw_form.ImageList3.ListImages(Asc(name) - 60).Picture, pcoord.X - 5, pcoord.Y - 5, 32, 32, 0, 0, 32, 32, &H990066
End If
End If
put_name_error:
End Sub
Public Sub set_depend_from_point(ByVal p%)
If m_poi(p%).data(0).from_wenti_no > 0 Then
Call C_display_wenti.set_m_depend_no(m_poi(p%).data(0).from_wenti_no)
If m_poi(p%).data(0).parent.element(0).ty = circle_ Then
   If m_Circ(m_poi(p%).data(0).parent.element(0).no).data(0).from_wenti_no > 0 Then
    Call C_display_wenti.set_m_depend_no(m_Circ(m_poi(p%).data(0).parent.element(0).no).data(0).from_wenti_no)
   End If
End If
End If
End Sub
Private Function regularity_coordinate(ByVal p%, coord As POINTAPI, out_coord As POINTAPI) As Boolean
Dim l%
t_coord = coord
If m_poi(p%).data(0).parent.element(0).no > 0 And m_poi(p%).data(0).parent.element(1).no > 0 Then
     out_coord = m_poi(p%).data(0).data0.coordinate
ElseIf m_poi(p%).data(0).parent.element(0).ty = line_ Then
   l% = m_poi(p%).data(0).parent.element(0).no
    Call orthofoot1(t_coord, _
     m_poi(m_lin(l%).data(0).data0.poi(0)).data(0).data0.coordinate, _
      m_poi(m_lin(l%).data(0).data0.poi(1)).data(0).data0.coordinate, _
        coord, 0, False)
   regularity_coordinate = True
ElseIf m_poi(p%).data(0).parent.element(0).ty = circle_ Then
   regularity_coordinate = True
Else
   out_coord = coord
   regularity_coordinate = True
End If
End Function
Public Sub change_point_by_line(ByVal l%)
'
Dim i%, tl%, tc%, w_no%
Dim p_coord As POINTAPI
For i% = 1 To last_conditions.last_cond(1).point_no
    w_no% = m_poi(i%).data(0).from_wenti_no
    If w_no% > 0 Then
    If C_display_wenti.m_no_(w_no%) = -1000 Then
    If m_poi(i%).data(0).degree = 1 Then
       If m_poi(i%).data(0).parent.element(0).ty = line_ And _
        m_poi(i%).data(0).parent.element(0).no = l% Then
       Else
       GoTo change_point_by_line_next
       End If
       
    ElseIf m_poi(i%).data(0).degree = 0 Then
       tl% = 0
       tc% = 0
       If m_poi(i%).data(0).parent.element(0).ty = line_ And _
           m_poi(i%).data(0).parent.element(0).no = l% Then
          If m_poi(i%).data(0).parent.element(1).ty = line_ Then
              tl% = m_poi(i%).data(0).parent.element(1).no
               Call inter_point_line_line(m_lin(tl%).data(0).data0.poi(0), _
                    m_lin(tl%).data(0).data0.poi(1), _
                    m_lin(l%).data(0).data0.poi(0), _
                    m_lin(l%).data(0).data0.poi(1), tl%, l%, i%, pointapi0, False, False, True)
          Else
              tc% = m_poi(i%).data(0).parent.element(1).no
          End If
       ElseIf m_poi(i%).data(0).parent.element(1).ty = line_ And _
           m_poi(i%).data(0).parent.element(1).no = l% Then
          If m_poi(i%).data(0).parent.element(0).ty = line_ Then
              tl% = m_poi(i%).data(0).parent.element(0).no
          Else
              tc% = m_poi(i%).data(0).parent.element(0).no
          End
       End If
     End If
End If
End If
End If
change_point_by_line_next:
Next i%
End Sub

Public Sub set_new_picture()
Dim i%
'for i%=1 to m_picutre.
End Sub

Attribute VB_Name = "print_wenti_form"
Option Explicit
Dim squre_root_level As Byte
Private Sub print_item(ob As Object, color As Byte, s As String, _
             ByVal display_or_delete As Byte)
Dim print_item_i%, n%
Dim ch As String
Dim sqr_no%
Dim ch1 As String * 1
  Call SetTextColor_(ob, color, display_or_delete)
If s = "" Or s = "1" Then
 Exit Sub
ElseIf InStr(3, s, "+", 0) > 0 Or InStr(3, s, "-", 0) > 0 Or _
        InStr(3, s, "#", 0) > 0 Or InStr(3, s, "@", 0) > 0 Then
 Call print_string(ob, color, s, 0, display_or_delete)
Else
n% = 1
For print_item_i% = 1 To Len(s)
 ch = Mid$(s, print_item_i%, 1)
  If print_item_i% < Len(s) Then
   ch1 = Mid$(s, print_item_i% + 1, 1)
  Else
   ch1 = "F"
  End If
If ch <> ch1 Then   '
   If n% > 1 Then
      If ch = "U" Then
         ch = "arrow\\u"
      ElseIf ch = "V" Then
         ch = "arrow\\v"
      End If
      Call display_char(ob, ch, middle, color, display_or_delete, False)
      Call display_char(ob, Trim(str(n%)), up, color, display_or_delete, False)
       n% = 1
   Else
    If ch = "^" Then
'     Call display_char(ob, Mid$(s, print_item_i% - 1, 1), middle, color, False, display_or_delete)
       print_item_i% = print_item_i% + 1
      Call display_char(ob, Mid$(s, print_item_i%, 1), up, color, display_or_delete, False)
       
    ElseIf ch = "[" Then 'Asc(ch) < 0 And ch <> LoadResString_(1461,"") And ch <> LoadResString_(856) And ch <> LoadResString_(1456,""811) Then
      Call print_sqr(ob, color, display_string_(read_sqr_no_from_string(s, print_item_i%, print_item_i%, ""), 0), _
                       display_or_delete)
    ElseIf ch = "_" Then
               Call display_char(ob, ch, middle, color, display_or_delete, False)

    '  icon_x = ob.CurrentX
    '   icon_y = ob.CurrentY
    '    ob.CurrentX = ob.CurrentX + 2
    '     ob.CurrentY = ob.CurrentY - 2
    ' If icon_display = 0 Then
    '  MDIForm1.Timer1.Enabled = True
    ' End If
    '  Call display_char(ob, set_display_string_for_set_value(ch, 0), middle, color, display_or_delete)
    Else
        If ch = "U" Then
          ch = "arrow\\u"
        ElseIf ch = "V" Then
          ch = "arrow\\v"
        End If
          Call display_char(ob, ch, middle, color, display_or_delete, False)
    End If
   End If
ElseIf ch = ch1 Then
   n% = n% + 1
End If
Next print_item_i%
End If
End Sub

Private Sub print_para(ob As Object, color As Byte, ByVal s As String, _
               ByVal display_or_delete As Byte, ty_ As Byte, item0 As String, _
                trs As Boolean)
                'ty_=0 第一项
Dim i%
Dim t%
Dim ty
Dim s0 As String
Dim s1 As String
Dim S2 As String
Dim s3 As String
Dim ch As String * 1
Dim ch1 As String
Dim ch2 As String
If item0 = "0" Then
  Exit Sub
End If
 ty = para_type(s, s0, s1, S2, s3)
  If s3 <> "" Then
    If item0 = "" Or item0 = "1" Or item0 = "_" Then
     Call print_para(ob, color, s1, display_or_delete, ty_, S2, trs)
     Call print_sqr(ob, color, S2, display_or_delete)
     Call print_para(ob, color, s3, display_or_delete, 1, item0, trs)
    Else
     Call display_char(ob, "(", middle, color, display_or_delete, False)
     Call print_string(ob, color, s0, ty_, display_or_delete)
     Call print_para(ob, color, s3, display_or_delete, 1, item0, trs)
     Call display_char(ob, ")", middle, color, display_or_delete, False)
    End If
    Exit Sub
  Else
   If S2 <> "1" Then
     Call print_para(ob, color, s1, display_or_delete, ty_, S2, trs)
     Call print_sqr(ob, color, S2, display_or_delete)
     Exit Sub
   End If
  End If
If s = "1" Or s = "#1" Then
 If ty_ = 0 Then
  If item0 <> "" And item0 <> "1" And item0 <> "_" Then
   Exit Sub
  Else
   Call display_char(ob, "1", middle, color, display_or_delete, False)
    Exit Sub
  End If
 Else
  If item0 <> "" And item0 <> "1" And item0 <> "_" Then
    Call display_char(ob, "+", middle, color, display_or_delete, False)
     Exit Sub
  Else
   Call display_char(ob, "+1", middle, color, display_or_delete, False)
  End If
 End If
ElseIf s = "-1" Or s = "@1" Then
  If item0 <> "" And item0 <> "1" And item0 <> "_" Then
   Call display_char(ob, "-", middle, color, display_or_delete, False)
    Exit Sub
  Else
   Call display_char(ob, "-1", middle, color, display_or_delete, False)
    Exit Sub
  End If
Else
  If Mid$(s, 1, 1) = "-" Or Mid$(s, 1, 1) = "@" Then
      Call display_char(ob, "-" + Mid$(s, 2, Len(s) - 1), middle, color, display_or_delete, False)
  ElseIf Mid$(s, 1, 1) = "+" Or Mid$(s, 1, 1) = "#" Then
      If ty_ = 0 Then
       Call display_char(ob, "+" + Mid$(s, 2, Len(s) - 1), middle, color, display_or_delete, False)
      Else
       Call display_char(ob, Mid$(s, 2, Len(s) - 1), middle, color, display_or_delete, False)
      End If
  Else
      If ty_ = 0 Then
       Call display_char(ob, s, middle, color, display_or_delete, False)
      Else
       Call display_char(ob, "+" + s, middle, color, display_or_delete, False)
      End If
  End If
  If item0 <> "" Or item0 <> "1" Then
  Else
  End If
End If
End Sub
Public Sub print_sqr(ob As Object, color As Byte, ByVal s As String, _
                                             ByVal display_or_delete As Byte)
Dim p1%, p2%, p3%, cuY%
Dim i%
Dim ch As String * 1
If s = "1" Or s = "" Then
Exit Sub
End If
cuY% = ob.CurrentY
 squre_root_level = squre_root_level + 1
If ob.ScaleMode = 3 Then
ob.CurrentY = ob.CurrentY + (squre_root_level * 2 - 2)
ElseIf ob.ScaleMode = 1 Then
ob.CurrentY = ob.CurrentY + (squre_root_level * 2 - 2) * 15
End If
p3% = ob.CurrentY ' +(squre_root_level * 2 - 2)
 Call display_char(ob, LoadResString_(1460, ""), middle, color, display_or_delete, False)  'ob.Print LoadResString_(1461,""); '根号
If ob.ScaleMode = 1 Then
   ob.CurrentX = ob.CurrentX - 70
End If
p1% = ob.CurrentX '
'******************************************************
Call remove_brace(s)
Call print_string(ob, color, s, 0, display_or_delete)
'ob.CurrentY = p3%
p2% = ob.CurrentX
'If Ob.ScaleMode = 3 Then
   If display_or_delete = 1 Then
    ob.Line (p2% - 2, p3% + 2)-(p1% - 2, p3% + 2), QBColor(color)
   ElseIf display_or_delete = 2 Then
    ob.Line (p2% - 2, p3% + 2)-(p1% - 2, p3% + 2), QBColor(7)
   Else
    ob.Line (p2% - 2, p3% + 2)-(p1% - 2, p3% + 2), QBColor(15)
   End If
'End If
ob.CurrentX = p2%
ob.CurrentY = cuY%
squre_root_level = squre_root_level - 1
End Sub
Public Sub print_string(ob As Object, color As Byte, ByVal s As String, _
                                ty_ As Byte, ByVal display_or_delete As Byte)                                    'ty=0 去括号
       'ty_=0 首次,
Dim t%, t1%, i%, b1%, b2%, ty%
Dim ts(1) As String
Dim s0 As String
Dim s1 As String
Dim S2 As String
Dim s3 As String
Dim ch1 As String
Dim ch2 As String
Dim s_type As Integer
Dim last_ch As String
Dim first_ch As String
Dim last_ch1 As String
If s = "" Then
 Exit Sub '空字符
End If
If remove_brace(s) Then '去括号
   Call display_char(ob, "(", middle, color, display_or_delete, False)
   Call print_string(ob, color, s, 0, display_or_delete)
   Call display_char(ob, ")", middle, color, display_or_delete, False)
    Exit Sub
Else
   s_type = string_type_(s, s0, s1, S2, s3)   '分类
   If s_type = 3 Then '分式
      Call print_string(ob, color, s1, 0, display_or_delete)
      Call display_char(ob, "/", middle, color, display_or_delete, False)
      Call print_string(ob, color, S2, 0, display_or_delete)
      Exit Sub
   ElseIf s_type = 2 Then 'fen
      Call print_string(ob, color, s1, 0, display_or_delete)
      Call print_string(ob, color, S2, 0, display_or_delete)
   Else
     If s3 <> "" Then
      Call print_string(ob, color, s0, ty_, display_or_delete)
      Call print_string(ob, color, s3, 1, display_or_delete)
      Exit Sub
     Else
      Call print_para(ob, color, s1, display_or_delete, ty_, S2, True)
      Call print_item(ob, color, S2, display_or_delete)
      Exit Sub
     End If
   End If
End If
End Sub
Public Function set_display_string_for_set_value(ch$, ty As Byte, is_depend As Boolean) As String
Dim i%
If ch$ = "" Then
 set_display_string_for_set_value = ""
  Exit Function
ElseIf Len(ch$) > 1 Then
  set_display_string_for_set_value = _
    set_display_string_for_set_value(Mid$(ch$, 1, 1), ty, is_depend) + _
     set_display_string_for_set_value(Mid$(ch$, 2, Len(ch$) - 1), ty, is_depend)
Else
If ch$ >= "a" And ch$ <= "w" Then
 For i% = 1 To last_used_char
  If used_char(i%).name = ch$ Then
   If used_char(i%).cond.ty = angle3_value_ Then
      set_display_string_for_set_value = _
         set_display_angle(angle3_value(used_char(i%).cond.no).data(0).data0.angle(0), is_depend)
   ElseIf used_char(i%).cond.ty = line_value_ Then
           set_display_string_for_set_value = _
         m_poi(line_value(used_char(i%).cond.no).data(0).data0.poi(0)).data(0).data0.name + _
         m_poi(line_value(used_char(i%).cond.no).data(0).data0.poi(1)).data(0).data0.name
   ElseIf used_char(i%).cond.ty = relation_ Then
      set_display_string_for_set_value = "(" + _
         m_poi(Drelation(used_char(i%).cond.no).data(0).data0.poi(0)).data(0).data0.name + _
         m_poi(Drelation(used_char(i%).cond.no).data(0).data0.poi(1)).data(0).data0.name + "/" + _
         m_poi(Drelation(used_char(i%).cond.no).data(0).data0.poi(2)).data(0).data0.name + _
         m_poi(Drelation(used_char(i%).cond.no).data(0).data0.poi(3)).data(0).data0.name + ")"
   Else
    set_display_string_for_set_value = ch$
     Exit Function
   End If
   If ty = 1 Then
     set_display_string_for_set_value = _
           "(" + set_display_string_for_set_value + ")"
   End If
   Exit Function
  End If
 Next i%
 set_display_string_for_set_value = ch$
Else
 set_display_string_for_set_value = ch$
End If
End If
End Function

Public Function is_relation_set_v(ByVal s1 As String, S2 As String, s3 As String, is_depend As Boolean) As Boolean
Dim ts1$, ts2$
Dim i%, j%
Dim ch$
For i% = 1 To Len(s1)
  ch$ = Mid$(s1, i%, 1)
   If ch$ >= "a" And ch$ <= "z" Then
      For j% = 1 To last_used_char
        If used_char(j%).name = ch$ Then
           If used_char(j%).cond.ty = relation_ Then
              ts1$ = ts1$ + m_poi(Drelation(used_char(j%).cond.no).data(0).data0.poi(0)).data(0).data0.name + _
                  m_poi(Drelation(used_char(j%).cond.no).data(0).data0.poi(1)).data(0).data0.name
              ts2$ = ts2$ + m_poi(Drelation(used_char(j%).cond.no).data(0).data0.poi(2)).data(0).data0.name + _
                  m_poi(Drelation(used_char(j%).cond.no).data(0).data0.poi(3)).data(0).data0.name
                   is_relation_set_v = True
                    GoTo is_relation_set_v_mark
           ElseIf used_char(j%).cond.ty = area_relation_ Then
              ts1$ = ts1$ + set_display_triangle(Darea_relation(used_char(j%).cond.no).data(0).area_element(0).no, is_depend, 0, 0)
              ts2$ = ts1$ + set_display_triangle(Darea_relation(used_char(j%).cond.no).data(0).area_element(1).no, is_depend, 0, 0)
                    GoTo is_relation_set_v_mark
           ElseIf used_char(j%).cond.ty = line_value_ Then
              ts1$ = ts1$ + m_poi(line_value(used_char(j%).cond.no).data(0).data0.poi(0)).data(0).data0.name + _
                  m_poi(line_value(used_char(j%).cond.no).data(0).data0.poi(1)).data(0).data0.name
                    GoTo is_relation_set_v_mark
           ElseIf used_char(j%).cond.ty = angle3_value_ Then
              ts1$ = ts1$ + m_poi(line_value(used_char(j%).cond.no).data(0).data0.poi(0)).data(0).data0.name + _
                  m_poi(line_value(used_char(j%).cond.no).data(0).data0.poi(1)).data(0).data0.name
                    GoTo is_relation_set_v_mark
           Else
              ts1$ = ts1$ + ch$
                    GoTo is_relation_set_v_mark
           End If
        End If
      Next j%
   ts1$ = ts1$ + ch$
is_relation_set_v_mark:
  Else
   ts1$ = ts1$ + ch$
   End If
Next i%
S2 = ts1$
s3 = ts2$
End Function

Public Sub SetTextColor_(ob As Object, color As Byte, display_or_delete As Byte)
    If display_or_delete = 1 Then '显示语句号
     If color = 0 Then
        color = condition_color
     End If
     Call SetTextColor(ob.hdc, QBColor(color))
    ElseIf display_or_delete = 2 Then
     Call SetTextColor(ob.hdc, QBColor(7))
    Else
     Call SetTextColor(ob.hdc, QBColor(15))   '消除
    End If
End Sub

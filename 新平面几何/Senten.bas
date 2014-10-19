Attribute VB_Name = "sentence"
Option Explicit
Global prove_step As Byte
Global input_char_no As Byte
Global inp As Integer
Global input_sentence_no(7, 15) As Integer
Type used_char_type
 name As String * 1
  cond As condition_type
End Type
Global used_char(30) As used_char_type '用过的小写字母
Global last_used_char As Integer
Public Function complete_input_sentence(ByVal n%) As Boolean
'判断输入是否完成
Dim i%, l%
Dim c$
Dim p1%, p2%
Dim s1 As String
Dim S2 As String
Dim s3 As String
Dim S4 As String
s1 = C_display_wenti.m_string(n%)
 p1% = InStr(1, s1, "!", 0)
  If p1% > 0 Then
   p2% = InStr(p1% + 1, s1, "~", 0)
    S2 = Mid$(s1, 1, p1% - 1)
     s3 = Mid$(s1, p1% + 1, p2% - p1% - 1)
      S4 = Mid$(s1, p2% + 1, Len(s1) - p2%)
   Else
   S2 = s1
   End If
l% = Len(S2)
 For i% = 1 To l%
  c$ = Mid$(S2, i%, 1)
   If c$ = "_" Or c$ = global_icon_char Then
    complete_input_sentence = False
     Exit Function
   End If
 Next i%
 l% = Len(s3)
 For i% = 1 To l%
  c$ = Mid$(s3, i%, 1)
   If c$ = "_" Or c$ = global_icon_char Then
    complete_input_sentence = False
     Exit Function
   End If
 Next i%
l% = Len(S4)
 For i% = 1 To l%
  c$ = Mid$(S4, i%, 1)
   If c$ = "_" Or c$ = global_icon_char Then
    complete_input_sentence = False
     Exit Function
   End If
 Next i%

  'Call from_sentence_to_input(n%)
  complete_input_sentence = True
  End Function
Public Function is_used_char(ByVal w%, ByVal n%) As Integer ', wenti_n%) As Byte
'　检查n% 　之前是否有c
Dim i%
 For i% = 1 To last_conditions.last_cond(1).point_no
 If C_display_wenti.m_point_no(w%, n%) = i% Then
 If m_poi(i%).data(0).data0.name = C_display_wenti.m_condition(w%, n%) Then
  is_used_char = i%
   Exit Function
  End If
  End If
 Next i%

is_used_char = 0
'***************************************************

End Function

Public Function get_input_info(ByVal n%, ByVal p%) As Boolean
Dim tp As Integer
tp = is_used_char(n%, p%)
'If tp = 0 Then
 'input_char_info = "此前尚未输入点" + c_display_wenti.m_condition(n%,p%) + "!"
  'get_input_info = True
'End If
Select Case C_display_wenti.m_no(n%)
Case -33
If p% = 2 And C_display_wenti.m_condition(n%, 2) = _
     C_display_wenti.m_condition(n%, 1) Then
 get_input_info = True
input_char_info = LoadResString_(1465, "")
 ElseIf p% = 3 And C_display_wenti.m_condition(n%, 3) = _
                        C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 ElseIf p% = 3 Then
 get_input_info = False
 End If
Case -32
If p% = 1 And C_display_wenti.m_condition(n%, 1) = _
                 C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
input_char_info = LoadResString_(1465, "")
 ElseIf p% = 2 And C_display_wenti.m_condition(n%, 2) = _
                 C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
input_char_info = LoadResString_(1465, "")

 ElseIf p% = 4 And C_display_wenti.m_condition(n%, 4) = _
                 C_display_wenti.m_condition(n%, 3) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 ElseIf p% = 5 Then
 get_input_info = False
 End If
Case -31 '
If p% = 1 And C_display_wenti.m_condition(n%, 1) = _
              C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
input_char_info = LoadResString_(1470, "")
 ElseIf p% = 2 Then
 If tp = 0 Then
If C_display_wenti.m_condition(n%, 2) = _
             C_display_wenti.m_condition(n%, 0) Or _
              C_display_wenti.m_condition(n%, 2) = _
               C_display_wenti.m_condition(n%, 1) Then
 get_input_info = True
input_char_info = LoadResString_(1470, "")
Else
  get_input_info = False
End If
 Else
   get_input_info = True
  input_char_info = LoadResString_(2110, "")
 End If

 ElseIf p% = 4 And C_display_wenti.m_condition(n%, 4) = _
                      C_display_wenti.m_condition(n%, 3) Then
 get_input_info = True
 input_char_info = LoadResString_(1140, "")
 ElseIf p% = 5 Then
 get_input_info = False
 End If
Case -30
If p% = 1 And C_display_wenti.m_condition(n%, 1) = _
                 C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
input_char_info = LoadResString_(1465, "")
 ElseIf p% = 2 And C_display_wenti.m_condition(n%, 2) = _
                C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
input_char_info = LoadResString_(1465, "")
ElseIf p% = 2 Then
If tp = 0 Then
 If C_display_wenti.m_condition(n%, 2) = _
           C_display_wenti.m_condition(n%, 1) Then
  get_input_info = True
   input_char_info = LoadResString_(1490, "")
 ElseIf C_display_wenti.m_condition(n%, 2) = _
              C_display_wenti.m_condition(n%, 0) Then
  get_input_info = True
   input_char_info = LoadResString_(1465, "")
 Else
  get_input_info = False
 End If
 Else
   get_input_info = True
  input_char_info = LoadResString_(2110, "")
 End If
 
 ElseIf p% = 4 And C_display_wenti.m_condition(n%, 4) = _
                     C_display_wenti.m_condition(n%, 3) Then
 get_input_info = True
 input_char_info = LoadResString_(1140, "")
 ElseIf p% = 5 Then
 get_input_info = False
 End If

Case -22, -23
'过□点平行□□的直线交□□于□
If p% = 2 And C_display_wenti.m_condition(n%, 2) = _
                      C_display_wenti.m_condition(n%, 1) Then
 get_input_info = True
input_char_info = LoadResString_(1470, "")
 ElseIf p% = 4 And C_display_wenti.m_condition(n%, 4) = _
                       C_display_wenti.m_condition(n%, 3) Then
 get_input_info = True
 input_char_info = LoadResString_(2115, "")
 ElseIf p% = 5 Then
 get_input_info = False
 End If


Case -20, -18, -17, -16
If p% = 1 And C_display_wenti.m_condition(n%, 1) = _
                  C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 ElseIf p% = 2 And (C_display_wenti.m_condition(n%, 2) = _
                          C_display_wenti.m_condition(n%, 1) _
                           Or C_display_wenti.m_condition(n%, 2) = _
                             C_display_wenti.m_condition(n%, 0)) Then
 get_input_info = True
 input_char_info = LoadResString_(2115, "")
 Else
 get_input_info = False
 End If

Case -19, -15, -14, -13, -12, -11, -10
If p% = 1 And C_display_wenti.m_condition(n%, 1) = _
                        C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 ElseIf p% = 2 And (C_display_wenti.m_condition(n%, 2) = _
                           C_display_wenti.m_condition(n%, 1) _
                           Or C_display_wenti.m_condition(n%, 2) = _
                              C_display_wenti.m_condition(n%, 0)) Then
 get_input_info = True
 input_char_info = LoadResString_(2120, "")
  ElseIf p% = 3 And (C_display_wenti.m_condition(n%, 3) = _
                         C_display_wenti.m_condition(n%, 2) _
                          Or C_display_wenti.m_condition(n%, 3) = _
                            C_display_wenti.m_condition(n%, 1) Or _
                               C_display_wenti.m_condition(n%, 3) = _
                                C_display_wenti.m_condition(n%, 0)) Then
 get_input_info = True
 input_char_info = LoadResString_(2120, "")

 Else
 get_input_info = False
 End If
Case -9
If p% = 1 And C_display_wenti.m_condition(n%, 1) = _
                    C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True

 input_char_info = LoadResString_(1470, "")
 ElseIf p% = 2 And (C_display_wenti.m_condition(n%, 2) = _
                          C_display_wenti.m_condition(n%, 1) _
                        Or C_display_wenti.m_condition(n%, 2) = _
                            C_display_wenti.m_condition(n%, 0)) Then
 get_input_info = True
 input_char_info = LoadResString_(2125, "")
  ElseIf p% = 3 And (C_display_wenti.m_condition(n%, 3) = _
                            C_display_wenti.m_condition(n%, 2) _
  Or C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 1) Or _
   C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 0)) Then
 get_input_info = True
 input_char_info = LoadResString_(2125, "")
   ElseIf p% = 4 And (C_display_wenti.m_condition(n%, 4) = C_display_wenti.m_condition(n%, 3) _
  Or C_display_wenti.m_condition(n%, 4) = C_display_wenti.m_condition(n%, 2) Or _
   C_display_wenti.m_condition(n%, 4) = C_display_wenti.m_condition(n%, 1) Or _
    C_display_wenti.m_condition(n%, 4) = C_display_wenti.m_condition(n%, 0)) Then
 get_input_info = True
 input_char_info = LoadResString_(2125, "")
Else
 get_input_info = False
End If
Case -8
If p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True

 input_char_info = LoadResString_(1470, "")
 ElseIf p% = 2 And (C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 1) _
  Or C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 0)) Then
 get_input_info = True
 input_char_info = LoadResString_(1485, "")
  ElseIf p% = 3 And (C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 2) _
  Or C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 1) Or _
   C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 0)) Then
 get_input_info = True
 input_char_info = LoadResString_(1485, "")
   ElseIf p% = 4 And (C_display_wenti.m_condition(n%, 4) = C_display_wenti.m_condition(n%, 3) _
  Or C_display_wenti.m_condition(n%, 4) = C_display_wenti.m_condition(n%, 2) Or _
   C_display_wenti.m_condition(n%, 4) = C_display_wenti.m_condition(n%, 1) Or _
    C_display_wenti.m_condition(n%, 4) = C_display_wenti.m_condition(n%, 0)) Then
 get_input_info = True
 input_char_info = LoadResString_(1485, "")
    ElseIf p% = 5 And (C_display_wenti.m_condition(n%, 5) = C_display_wenti.m_condition(n%, 4) _
  Or C_display_wenti.m_condition(n%, 5) = C_display_wenti.m_condition(n%, 3) Or _
   C_display_wenti.m_condition(n%, 5) = C_display_wenti.m_condition(n%, 2) Or _
    C_display_wenti.m_condition(n%, 5) = C_display_wenti.m_condition(n%, 1) Or _
     C_display_wenti.m_condition(n%, 5) = C_display_wenti.m_condition(n%, 0)) Then
 get_input_info = True
 input_char_info = LoadResString_(1485, "")

Else
 get_input_info = False
End If
Case -7
If p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 ElseIf p% = 3 And C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 2) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 ElseIf p% > 3 Then
 get_input_info = False
 End If
Case -6
If p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
Else
 input_char_info = False
End If
Case -5
If p% > 3 Then
 get_input_info = False
End If
Case -4
If p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 ElseIf p% = 2 And C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 1) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 ElseIf p% = 4 And C_display_wenti.m_condition(n%, 4) = C_display_wenti.m_condition(n%, 3) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 ElseIf p% = 5 And C_display_wenti.m_condition(n%, 5) = C_display_wenti.m_condition(n%, 4) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 End If
Case -3
If p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
input_char_info = LoadResString_(1465, "")
 ElseIf p% = 2 And C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
input_char_info = LoadResString_(1465, "")

 ElseIf p% = 4 And C_display_wenti.m_condition(n%, 4) = C_display_wenti.m_condition(n%, 3) Then
 get_input_info = True
 input_char_info = LoadResString_(1465, "")
 ElseIf p% = 5 Then
 get_input_info = False
 End If


Case -2

If p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
input_char_info = LoadResString_(1465, "")
 ElseIf p% = 3 And C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 2) Then
 get_input_info = True
input_char_info = LoadResString_(1465, "")

 ElseIf p% = 5 And C_display_wenti.m_condition(n%, 5) = C_display_wenti.m_condition(n%, 4) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 ElseIf p% = 4 Or p% = 5 Then
 get_input_info = False
 End If

Case -1
If p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 ElseIf p% = 3 And C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 2) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 End If

'inpcond(0) =直线□□上任取一点□
Case 0
If tp > 0 Then
 get_input_info = True
 input_char_info = LoadResString_(1490, "")
Else
get_input_info = False
End If
'inpcond(1) = 直线□□上任取一点□    '16
Case 1, 4, 5
 If p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 ElseIf p% = 2 Then
  If tp > 0 Then
  get_input_info = True
 input_char_info = LoadResString_(1490, "") '"重名，请重新给点命名。"
 Else
 get_input_info = False
 End If
End If
 'inpcond(2) = "在过点□且平行□□的直线上任取一点□"  '17
Case 2 '平行线
 If p% = 1 And C_display_wenti.m_condition(n%, 0) = C_display_wenti.m_condition(n%, 1) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "") '端点重合
 ElseIf p% = 2 And (C_display_wenti.m_condition(n%, 0) = C_display_wenti.m_condition(n%, 2) Or _
    C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 2)) Then
  get_input_info = True
   input_char_info = "平行线重合"
 ElseIf p% = 3 And C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 3) Then
  get_input_info = True
   input_char_info = LoadResString_(1470, "") '端点重合
 Else
  get_input_info = False
 End If
'inpcond(3) = "在过点□且垂直□□的直线上任取一点□" '18
'inpcond(4) = 在□□的垂直平分线上任取一点□  '19
'inpcond(5) = 取线段□□的中点□  '22
'inpcond(6) = □是线段□□上分比为_的分点 '15
Case 6
 If (p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0)) Or _
 (p% = 2 And C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 0)) Then
 get_input_info = True
 input_char_info = LoadResString_(2135, "")
 ElseIf p% = 2 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 2) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 
 ElseIf p% = 0 Then
   If tp > 0 Then
  get_input_info = True
 input_char_info = LoadResString_(1490, "")
 Else
  get_input_info = False
 End If

 End If
'inpcond(7) = ⊙□(_)上任取一点□  '2
Case 7
 If p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
 input_char_info = LoadResString_(2140, "")
 ElseIf p% = 2 Then
  If tp > 0 Then
  get_input_info = True
 input_char_info = LoadResString_(1490, "")
 Else
 get_input_info = False
 End If
End If
Case 8
 If (p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0)) Or _
 (p% = 2 And (C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 0) Or _
  C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 1))) Then
 get_input_info = True
 input_char_info = LoadResString_(1510, "")
  ElseIf p% = 3 Then
  If tp > 0 Then
  get_input_info = True
 input_char_info = LoadResString_(1490, "")
 Else
 get_input_info = False
 End If
End If

'inpcond(9) = 直线□□和直线□□交于点□     '10
Case 9
 If (p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0)) Or _
 (p% = 3 And C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 3)) Then
 get_input_info = True
 input_char_info = LoadResString_(1470, "")
 ElseIf p% = 4 Then
  If tp > 0 Then
  get_input_info = True
 input_char_info = LoadResString_(1490, "")
  Else
 get_input_info = False
 End If

End If

'inpcond(10) = "过□垂直□□的直线交⊙□于□" '11
Case 10, 16
 If p% = 2 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 2) Then
 get_input_info = True
 input_char_info = LoadResString_(1495, "")
 ElseIf p% = 4 And C_display_wenti.m_condition(n%, 4) = C_display_wenti.m_condition(n%, 3) Then
 get_input_info = True
 input_char_info = LoadResString_(2140, "")
 ElseIf p% = 5 Then
If tp > 0 Then
  get_input_info = True
 input_char_info = LoadResString_(1490, "")
  Else
 get_input_info = False
 End If

End If

'inpcond(11) = □是直线□□与⊙□(_)的一个交点  '12
Case 11
 If p% = 2 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 2) Then
 get_input_info = True
 input_char_info = LoadResString_(1495, "")
  ElseIf p% = 4 And C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 4) Then
 get_input_info = True
 input_char_info = LoadResString_(2140, "")
 ElseIf p% = 0 Then
 If tp > 0 Then
  get_input_info = True
  input_char_info = LoadResString_(1490, "")
  Else
 get_input_info = False
 End If

End If

'inpcond(12) = "⊙□_和⊙□_相交于点□取另一个交点□"  '13
Case 12
If p% = 2 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 2) Then
 get_input_info = True
 input_char_info = LoadResString_(2140, "")
  ElseIf p% = 2 And C_display_wenti.m_condition(n%, 0) = C_display_wenti.m_condition(n%, 2) Then
 get_input_info = True
 input_char_info = LoadResString_(455, "")
 ElseIf p% = 3 And C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 3) Then
 get_input_info = True
 input_char_info = LoadResString_(2140, "")
  ElseIf p% = 4 And (C_display_wenti.m_condition(n%, 0) = C_display_wenti.m_condition(n%, 4) Or _
   C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 4)) Then
 get_input_info = True
 input_char_info = LoadResString_(460, "")
 ElseIf p% = 5 Then
 If tp > 0 Then
  get_input_info = True
 input_char_info = LoadResString_(1490, "")
  Else
 get_input_info = False
 End If

End If
'inpcond(13) = □是⊙□(_)和⊙□(_)的一个交点   '14
Case 13
 If p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0) Then
 get_input_info = True
 input_char_info = LoadResString_(2140, "")
 ElseIf p% = 3 And C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 3) Then
 get_input_info = True
 input_char_info = LoadResString_(2140, "")
 ElseIf p% = 0 Then
 If tp > 0 Then
  get_input_info = True
 input_char_info = LoadResString_(1490, "")
  Else
 get_input_info = False
 End If

End If



'inpcond(14) =过□作直线□□的垂线垂足为□
Case 14
 If p% = 2 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 2) Then
 get_input_info = True
 input_char_info = LoadResString_(1495, "")
 ElseIf p% = 3 Then
 If tp > 0 Then
  get_input_info = True
 input_char_info = LoadResString_(1490, "")
  Else
 get_input_info = False
 End If

End If
'inpcond(15) = "过□平行□□的直线交□□于□"
'inpcond(16) = "过□平行□□的直线和过□垂直□□的直线交于点□"
'inpcond(17) = "过□垂直□□的直线交□□于□"
Case 15, 17
 If p% = 2 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 2) Then
 get_input_info = True
 input_char_info = LoadResString_(1495, "")
 ElseIf p% = 4 And C_display_wenti.m_condition(n%, 4) = C_display_wenti.m_condition(n%, 3) Then
 get_input_info = True
 input_char_info = LoadResString_(1495, "")
     ElseIf p% = 5 Then
     If tp > 0 Then
  get_input_info = True
 input_char_info = "重名，请重新给点命名。"
  Else
 get_input_info = False
 End If

End If
'inpcond(18) =□是△□□□的重心    '4
'inpcond(19) =□是△□□□的外接圆的圆心 '7
'inpcond(20) =□是△□□□的垂心 '27
'inpcond(21) =□是△□□□的内切圆的圆心  '29
Case 18, 19, 20, 21
If (p% = 2 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 2)) Or _
(p% = 3 And (C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 1) Or _
 C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 3))) Then
  get_input_info = True
 input_char_info = LoadResString_(1500, "")
ElseIf (p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0)) Or _
       (p% = 2 And C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 0)) Or _
       (p% = 3 And C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 0)) Then
       get_input_info = True
          input_char_info = "重名，请重新给点命名。"
ElseIf p% = 0 Then
If tp > 0 Then
  get_input_info = True
 input_char_info = "重名，请重新给点命名。"
 Else
 get_input_info = False
 End If
End If
'inpcond(22) = "□是既约多项式_的零点"  '6
'inpcond(23) = □、□、□、□四点共圆
Case 23
If (p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0)) Or _
 (p% = 2 And (C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 0) Or _
  C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 1))) Or _
   (p% = 3 And (C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 0) Or _
     C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 1) Or _
      C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 2))) Then
   get_input_info = True
 input_char_info = LoadResString_(1505, "")
End If
'inpcond(24) = □、□、□三点共线
Case 24
If (p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0)) Or _
 (p% = 2 And (C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 0) Or _
  C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 1))) Then
   get_input_info = True
 input_char_info = LoadResString_(1510, "")
End If
'inpcond(25) = "线段□□和□□长相等，即｜□□｜＝｜□□｜"
Case 25, 27, 28
If (p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0)) Or _
(p% = 3 And C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 2)) Then
    get_input_info = True
 input_char_info = LoadResString_(1515, "")
ElseIf p% = 3 And (C_display_wenti.m_condition(n%, 0) = C_display_wenti.m_condition(n%, 2) And _
 C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 3)) Or _
  (C_display_wenti.m_condition(n%, 0) = C_display_wenti.m_condition(n%, 3) And _
   C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 2)) Then
    get_input_info = True
 If C_display_wenti.m_no(n%) = 25 Then
 input_char_info = LoadResString_(1520, "")
 ElseIf C_display_wenti.m_no(n%) = 27 Then
 input_char_info = LoadResString_(1525, "")
  Else
   input_char_info = LoadResString_(1530, "")

 End If
End If
   'inpcond(26) = 点□是线段□□的中点
Case 26, 29
If (p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0)) Or _
(p% = 2 And C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 0)) Then
     get_input_info = True
 If C_display_wenti.m_no(n%) = 26 Then
 input_char_info = LoadResString_(2145, "")
 Else
 input_char_info = LoadResString_(2150, "")
 End If
ElseIf p% = 2 And C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 1) Then
     get_input_info = True
 input_char_info = LoadResString_(1515, "")
End If
'inpcond(27) = □□∥□□
'inpcond(28) = □□⊥□□

'inpcond(29) = 点□位于线段□□的垂直平分线上
'inpcond(30) = loadresstring_(289)
Case 30
If p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0) Then
     get_input_info = True
 input_char_info = LoadResString_(1515, "")
ElseIf p% = 2 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 2) Then
     get_input_info = True
 input_char_info = LoadResString_(1515, "")
ElseIf p% = 4 And C_display_wenti.m_condition(n%, 4) = C_display_wenti.m_condition(n%, 3) Then
     get_input_info = True
input_char_info = LoadResString_(1515, "")
ElseIf p% = 5 And C_display_wenti.m_condition(n%, 5) = C_display_wenti.m_condition(n%, 4) Then
     get_input_info = True
input_char_info = LoadResString_(1515, "")
End If
'inpcond(31) = "线段□□上的分点□满足□□：□□＝_"
Case 31
If p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0) Then
     get_input_info = True
 input_char_info = LoadResString_(1515, "")
ElseIf p% = 3 And C_display_wenti.m_condition(n%, 2) = C_display_wenti.m_condition(n%, 3) Then

     get_input_info = True
 input_char_info = LoadResString_(1470, "")
End If
'inpcond(32) = "四条线段成比例,即□□：□□＝□□：□□"
Case 32
If (p% = 1 And C_display_wenti.m_condition(n%, 1) = C_display_wenti.m_condition(n%, 0)) Or _
 (p% = 3 And C_display_wenti.m_condition(n%, 3) = C_display_wenti.m_condition(n%, 2)) Or _
(p% = 5 And C_display_wenti.m_condition(n%, 5) = C_display_wenti.m_condition(n%, 4)) Or _
(p% = 7 And C_display_wenti.m_condition(n%, 7) = C_display_wenti.m_condition(n%, 6)) Then
     get_input_info = True
 input_char_info = LoadResString_(1515, "")
End If
End Select
End Function

Public Sub input_sentence_y(ByVal i%, ByVal inp_%, need_draw As Byte)
Dim t_no%
If run_type > 0 Then
 Exit Sub
End If
inp = inp_%
MDIForm1.Inputcond.Enabled = False
MDIForm1.conclusion.Enabled = False
MDIForm1.Toolbar1.Buttons(21).Image = 34
'Wenti_form.Picture1.CurrentY = display_wenti_v_position%
If event_statue = wait_for_modify_sentence Then
 If list_type_for_input = input_condition_statue Then
 ElseIf list_type_for_input = input_prove_by_hand Or _
   input_type = input_add_point Then
    Call C_display_wenti.set_m_ty(modify_wenti_no, 4)
     'Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, 0, modify_wenti_no, _
        modify_wenti_no + 3, 0, 0, 0)
 
 End If
 Call input_sentence(1, modify_wenti_no)  ', modify_wenti_no, True)
ElseIf event_statue = wait_for_input_condition Then
modify_wenti_no = 0
     Call C_display_wenti.set_m_no(0, inp, modify_wenti_no)  '输入问题类型
      'modify_wenti_no = C_display_wenti.m_last_input_wenti_no '记录正在操作的输入语句号
 If inp < 23 Then '输入语句（已知条件）
 Else '结论
  If set_or_prove = 0 Then
   set_or_prove = 1
  End If
  'If wenti_type = 0 Then
  'call C_display_wenti.m_display_string.item(wenti_no).set_m_ty = _
   set_conclusion_ty(C_display_wenti.m_display_string.item(wenti_no).no)
    ' 记录结论的类型
 End If
  'Call C_display_wenti.input_m_sentences(Wenti_form.Picture1, 1, C_display_wenti.m_last_input_wenti_no, _
                     C_display_wenti.m_last_input_wenti_no, 1, False, 0)
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
ElseIf event_statue = input_prove_by_hand Or _
   event_statue = input_add_point Then
     'If record0.condition_type(1) > 0 And i% = 3 Then
      'prove_type = record0.condition_type(1) - 1
       'Call display_prove_inform(modify_wenti_no, delete)
      'If prove_type <> 3 And prove_type <> 2 Then
       'Call display_input_condi(Wenti_form.Picture1, delete, modify_wenti_no, modify_wenti_no + 3, 0)
       'Call init_input_cond(modify_wenti_no)
      'wenti_no = wenti_no - 1
     'End If
     'End If
     
     modify_wenti_no = C_display_wenti.m_last_input_wenti_no
      Call C_display_wenti.set_m_no(0, C_display_wenti.m_last_input_wenti_no, inp)
       Call C_display_wenti.set_m_ty(C_display_wenti.m_last_input_wenti_no, 4)
        ' Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, 1, _
          C_display_wenti.m_last_input_wenti_noo, C_display_wenti.m_last_input_wenti_no + 3, 0, 0, 0)
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
End If
 event_statue = wait_for_input_char
  'Call get_input_char
  If inp <> 38 And inp <> 50 Then
  Wenti_form.Picture1.SetFocus
  End If
End Sub

Public Sub get_input_char() '询问点的信息
 If event_statue = wait_for_input_char Then
   While event_statue = wait_for_input_char '等待事件发生
    DoEvents
   Wend
  If GetAsyncKeyState(&H27) = 1 Then
   Call C_display_wenti.m_input_char(Wenti_form.Picture1, "?")
  End If
 End If
End Sub

Public Function get_new_char(ByVal p%) As String
If p% > 0 And p% < 95 Then
If m_poi(p%).data(0).data0.name = "" Or m_poi(p%).data(0).data0.name = empty_char Then
    Call set_point_name(p%, next_char(p%, "", 0, 0)) ' next_char(p%)
End If
    get_new_char = m_poi(p%).data(0).data0.name
End If
End Function
Public Sub delete_name(ByVal ch$)
Dim j%, k%, l%
  If C_display_wenti.m_last_input_wenti_no > 0 Then
  For j% = C_display_wenti.m_last_input_wenti_no To 1 Step -1   'While j% < wenti_no
   If C_display_wenti.m_no(j%) = 0 Then
    For k% = 0 To 6
     If C_display_wenti.m_condition(j%, k%) = ch$ Then
     'Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, 0, j%, _
         condition_no, 0, 0, 0)
      For l% = k% To 18
       Call C_display_wenti.set_m_condition(j%, _
                   C_display_wenti.m_condition(j%, j% + 1), j%)
       Call C_display_wenti.m_point_no(j%, _
                      C_display_wenti.m_point_no(j%, j% + 1))
      Next l%
     End If
    Next k%
    ElseIf temp_wenti_cond(j%).no = 8 Then
    'For k% = 0 To 19
     If C_display_wenti.m_condition(j%, 2) = ch$ Then
    ' Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, 0, j%, _
                                  condition_no, 0, 0, 0)
     'Call set_wenti_cond_1(C_display_wenti.m_point_no(j%,0), _
                            C_display_wenti.m_point_no(j%,3), _
                             C_display_wenti.m_point_no(j%,1), _
                              C_display_wenti.m_point_no(3), _
                                C_display_wenti.m_point_no(8), _
                                C_display_wenti.m_point_no(8), _
                                 C_display_wenti.m_point_no(8), _
                                  C_display_wenti.m_point_no(8) , _
                                   C_display_wenti.m_point_no(8), _
                                    C_display_wenti.m_point_no(8))
     End If
   Else
    For k% = 0 To 50
     If C_display_wenti.m_condition(j%, k%) = ch$ Then
      For l% = j% To C_display_wenti.m_last_input_wenti_no - 2
         Call C_display_wenti.Remove_m_string(l%)
      Next l%
       j% = j% - 1
        Call init_wenti(C_display_wenti.m_last_input_wenti_no)
         If C_display_wenti.m_last_input_wenti_no = 0 Then
          GoTo delete_name_mark1
         End If
     End If
   Next k%
   End If
Next j%
delete_name_mark1:
End If
End Sub
Public Function is_used_char0(ByVal c As String) As Boolean
Dim i%
For i% = 1 To last_conditions.last_cond(1).point_no
 If m_poi(i%).data(0).data0.name = c Then
  is_used_char0 = True
   Exit Function
 End If
Next i%
End Function
'*******************************************************************
Public Sub input_sentence(ByVal display_or_delete As Byte, ByVal n%)
Dim temp_input_type As Boolean
Dim i%, l%
Dim complete As Boolean
'Wenti_form.Picture1.CurrentY = display_wenti_v_position%
If event_statue <> ready And _
           event_statue <> wait_for_modify_sentence Then
   '其他操作未完成
 Exit Sub
ElseIf event_statue = ready Then
 event_statue = wait_for_input_sentence
 '进入输入
End If

input_char_no = 0
If list_type_for_input = input_condition_statue Then
 If inp < 23 And n% > last_condition Then
'　输入结论状态下输入条件
  Exit Sub
 End If

temp_input_type = input_type  '纪录输入类，已知，结论?????


'If event_statue = wait_for_input_sentence Then
 '前一句已输完,进入输入语句状态
 
 '非修改，非输入字符状态
 '***************************************************************
 If n% = 0 And inp < 23 Then
             '首次输入，输入条件
     Call SetTextColor(Wenti_form.Picture1.hdc, 0)
    Wenti_form.Picture1.CurrentX = display_wenti_h_position%
      Wenti_form.Picture1.Print LoadResString_(2155, "") + ":"
       'Call C_display_wenti.m_display_string.item(n%).set_m_ty(1)
 ElseIf C_display_wenti.m_last_input_wenti_no = last_condition And inp > 22 Then
        Wenti_form.Picture1.CurrentY = 20 * (n% + 1)
         Wenti_form.Picture1.CurrentX = 0
          'Call C_display_wenti.m_display_string.item(n%).set_m_ty(1)
           Call SetTextColor(Wenti_form.Picture1.hdc, 0)
  If inp < 35 Then
    'Call C_display_wenti.m_display_string.item(n%).set_m_ty(2)
     Call SetTextColor(Wenti_form.Picture1.hdc, 0)
    Wenti_form.Picture1.CurrentX = display_wenti_h_position%
     solve_problem_type = 0
      input_p1% = Wenti_form.Picture1.CurrentY
       Wenti_form.Picture1.Print LoadResString_(465, "")
         problem_type = False
  Else
     Call SetTextColor(Wenti_form.Picture1.hdc, 0)
    Wenti_form.Picture1.CurrentX = display_wenti_h_position%
    solve_problem_type = 2
     Wenti_form.Picture1.Print LoadResString_(440, "")
      'Call C_display_wenti.m_display_string.item(n%).set_m_ty(3)
  problem_type = True
  End If
      End If
ElseIf list_type_for_input = input_prove_by_hand Or _
 list_type_for_input = input_add_point Then
If prove_step = 0 Then
        Wenti_form.Picture1.CurrentY = 20 * (n% + 2)
         Wenti_form.Picture1.CurrentX = 0
          Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(0))

 If problem_type = False Then
   Wenti_form.Picture1.CurrentX = display_wenti_h_position%
    Wenti_form.Picture1.Print LoadResString_(450, "")
 Else
   Wenti_form.Picture1.CurrentX = display_wenti_h_position%
    Wenti_form.Picture1.Print LoadResString_(445, "")
 End If
 prove_step = 1
End If
End If
      
  '****************************************************************
           Call C_display_wenti.set_m_no(0, n%, inp)
            Call C_display_wenti.set_m_string(n%, "", inpcond(inp).inpcond, "", "", _
                  "", 0, 0, 0)
             'wenti_cond(n%).no = inp
              '初始化输入语句
    If list_type_for_input = input_condition_statue Then
    If inp < 23 Then '显示　条件
     'Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, display_or_delete, _
        n%, n% + 1, False, 0, 0)
     Else   'If input_type=input_condition_statue Then
     'Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, display_or_delete, _
        n%, n% + 2, False, 0, 0)
     End If
     ElseIf list_type_for_input = input_prove_by_hand Or _
     list_type_for_input = input_add_point Then
     'Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, display_or_delete, _
            n%, n% + 3, False, 0, 0)
     End If
       '计输入语句号
        '进入输入字符状态
'***********************************************************************
 'If event_statue = wait_for_modify_sentence Then
 'For i% = 0 To 50
 ' wenti_cond(n%).m_condition(i%) = empty_char
 ' wenti_cond(n%).point_no(i%) = -1
 ' Next i%
 'End If
     '修改
'     If (modify_wenti_no < last_condition And _
'         inp > 22) Or _
'          (modify_wenti_no >= last_condition And inp > 23) Then
'      Wenti_form.Text1.Visible = 1   '显示警告
'      Exit Sub
'      End If
         '插入新语句
'***************************************************************
         '*****************************************************
'         If modify_wenti_no < last_condition _
'           And wenti_no > last_condition Then
'            For i = wenti_no - 1 To last_condition Step -1 '移动结论
'             Call display_input_condi(display, i, i + 3, 0)
'              Call display_input_condi(delete, i, i + 2, 0)
'            Next i
              '*****************************************
'               wenti_form.CurrentX = 0
'                wenti_form.CurrentY = (last_condition + 1) * 20
'                 Call settextcolor(wenti_form.hDC, QBColor(15))
'                wenti_form.Print loadresstring_(1044)
'                 Call settextcolor(wenti_form.hDC, QBColor(0))
'                wenti_form.Print loadresstring_(1044)
'           '***********************************************
'            input_type = input_condition_statue                     '移动条件
'            For i = last_condition - 1 To modify_wenti_no Step -1
'              Call display_input_condi(display, i, i + 2, 0)
'               Call display_input_condi(delete, i, i + 1, 0)
 
'            Next i
       
         '****************************************************************
'      Else              '修改结论
'             For i = wenti_no - 1 To modify_wenti_no Step -1
'              If input_type = input_condition_statue Then
'               Call display_input_condi(display, i, i + 2, 0)
'               Call display_input_condi(delete, i, i + 1, 0)

'               ElseIf input_type = input_conclusion Then
'               Call display_input_condi(display, i, i + 3, 0)
'               Call display_input_condi(delete, i, i + 2, 0)
'               End If
'           Next i
'      End If
              '**********************************
'              For i = wenti_no - 1 To modify_wenti_no Step -1
'               wenti_cond(i + 1) = wenti_cond(i)
'              display_input_condition(i + 1) = display_input_condition(i)
'              Next i
 '************************************
             
'            input_type = temp_input_type
           'wenti_cond(modify_wenti_no).no = inp '语句号存入
'            display_input_condition(modify_wenti_no).no = inp
'            display_input_condition(modify_wenti_no).cond = inpcond(inp)
'         If input_type = input_condition_statue Then
'          Call display_input_condi(display, modify_wenti_no, modify_wenti_no + 1, 0)

'         Else   'If input_type=input_condition_statue Then
'       Call display_input_condi(display, modify_wenti_no, modify_wenti_no + 2, 0)
         
'         End If
         
 '        event_statue = wait_for_modify_char
 'End If
out:
                    modify_statue = no_modify
 If list_type_for_input = input_condition_statue Then
         If inp < 23 And event_statue = wait_for_input_sentence Then 'type = input_condition_statue Then
           last_condition = last_condition + 1
         End If
 End If
         If modify_wenti_no = C_display_wenti.m_last_input_wenti_no Then
operate_step(C_display_wenti.m_last_input_wenti_no).last_point = last_conditions.last_cond(1).point_no
         'modify_wenti_no = modify_wenti_no + 1
         End If
     
'**************************************************************************
       'If wenti_form.Height < 20 * wenti_no + 40 Then
       '  wenti_form.Height = wenti_form.Height + 20
       '  wenti_form.Top = wenti_form.Top - 20
       'End If         '升高
    
 '*****************************************************************************

      event_statue = wait_for_input_char

End Sub



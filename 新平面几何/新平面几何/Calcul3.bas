Attribute VB_Name = "CALCUL3"
Option Explicit
Global para_begin As Boolean
 Global string_begin As Boolean
Global is_multi_item As Boolean
 Global is_multi_para As Boolean
 Global is_divide As Boolean
  Dim temp_int1(8, 8) As Integer
   Dim temp_int2(8, 8) As Integer
 Type para_item_type
  item() As String
   it() As String
    pA() As String
     last_it As Integer
 End Type
' Global Const polynomial = 0
' Global Const squre_root_ = 1
' Global Const ln_ = 2
' Global Const exp_ = 3
' Global Const sin_ = 4
' Global Const cos_ = 5
' Global Const tan_ = 6
' Global Const ctan_ = 7
' Global Const Asin_ = 8
' Global Const Acos_ = 9
' Global Const Atan_ = 10
' Global Const Actan_ = 11
  
' Type element_string_type '基本表达式
'  ty As Byte
'  parameter As String
'  str_v As String
' End Type
 Global number_string() As String
 Global last_number_string As Integer
Public Function min_for_long(ByVal A&, ByVal b&) As Long
If A& < b& Then
min_for_long = A&
Else
min_for_long = b&
End If
End Function

Public Function max_for_long(ByVal A&, ByVal b&) As Long
If A& > b& Then
max_for_long = A&
Else
max_for_long = b&
End If
End Function
Public Function l_gcd(ByVal A As Long, ByVal b As Long) As Long '求两个整数的最大公约数，输入A，B 输出 1_gcd
Dim c As Long
A = Abs(A)
 b = Abs(b)
If A = 0 Then
 l_gcd = b
ElseIf b = 0 Then
l_gcd = A
Else
c = A Mod b
l_gcd = l_gcd(b, c)
End If
End Function
Function gcd(ByVal A As Long, ByVal b As Long, out_A As Long, out_b As Long) As Long '求两个整数的最大公约数，输入A，B 输出 gcd 最大公约数，out_A_,out_B,A,B除以gcd的商
Dim c As Long
Dim t_A(1) As Long
t_A(0) = A
t_A(1) = b
A = Abs(A)
 b = Abs(b)
If A = 0 Then
 gcd = b
ElseIf b = 0 Then
gcd = A
Else
c = A Mod b
gcd = gcd(b, c, 0, 0)
End If
out_A = t_A(0) / gcd
out_b = t_A(1) / gcd
End Function

Function max(ByVal A As Integer, ByVal b As Integer) As Integer  '求出最大值
If A > b Then
max = A
Else
max = b
End If
End Function

Function min(A As Integer, b As Integer) As Integer '求出最小值
If A > b Then
min = b
Else
min = A
End If
End Function
Public Function add_brace(ByVal s As String, ty As String) As String
Dim i%, j%, l%, m%, n%, n1%, n2%
Dim ts$
Dim sg As String
Dim ts1$
l% = Len(s)
add_brace = s
 If l% < 3 Then
  Exit Function
 End If
 If Mid$(s, 1, 1) = "(" And Mid$(s, Len(s), 1) = ")" Then
  If is_brace(s, 1, Len(s)) Then
   add_brace = s
   Exit Function
  End If
 End If
 ts$ = s
     n% = InStr(1, ts, "#")
     m% = InStr(1, ts, "@")
     i% = InStr(1, ts, "+")
     j% = InStr(1, ts, "-")
 l% = InStr(1, ts, "(")
   If l% > 0 Then
    Call read_brace(ts, l%, n1%, n2%)
     If i% < 2 And j% < 2 Then
      Exit Function
     ElseIf i% < l% Or j% < l% Then
      add_brace = "(" + ts + ")"
     Else
      ts1 = Mid$(ts, n2% + 1, Len(ts) - n2%)
       i% = InStr(1, ts1, "+")
        j% = InStr(1, ts1, "-")
         If i% > 0 Or j% > 0 Then
          add_brace = "(" + ts + ")"
         End If
     End If
Else
  If n% = 1 Then
     n% = max(InStr(2, ts, "#"), n%)
  End If
  If m% = 1 Then
     m% = max(InStr(2, ts, "@"), m%)
  End If
  If i% = 1 Then
  i% = max(InStr(2, ts, "+"), i%)
  End If
  If j% = 1 Then
  j% = max(InStr(2, ts, "-"), j%)
  End If
 If i% > 1 Or j% > 1 Or m% > 1 Or n% > 1 Then
          add_brace = "(" + ts + ")"
 End If
End If
End Function
Private Function divide_item(ByVal i1$, ByVal i2$, _
     oI1$, oI2$) As String
Dim i%, j%, n1%, n2%
On Error GoTo divide_item_error
If Mid$(i1$, 1, 1) = "(" Then
 i1$ = Mid$(i1$, 1, Len(i1$) - 1)
  i1$ = Mid$(i1$, 2, Len(i1$) - 1)
End If
If Mid$(i2$, 1, 1) = "(" Then
 i2$ = Mid$(i2$, 1, Len(i2$) - 1)
  i2$ = Mid$(i2$, 2, Len(i2$) - 1)
End If
If i2$ = "1" Then
divide_item = i1$
 oI1$ = i1$
  oI2$ = i2$
   Exit Function
End If
n1% = InStr(1, i1$, "/", 0)
n2% = InStr(1, i2$, "/", 0)
If n1% = 0 And n2% = 0 Then
i% = 1
Do While i% <= Len(i1$)
 j% = 1
  Do While j% <= Len(i2$)
If Mid$(i1$, i%, 1) = Mid$(i2$, j%, 1) Then
  i1$ = Mid$(i1$, 1, i% - 1) + Mid$(i1$, i% + 1, Len(i1$) - i%)
   i2$ = Mid$(i2$, 1, j% - 1) + Mid$(i2$, j% + 1, Len(i2$) - j%)
  GoTo divide_item_mark0
Else
 j% = j% + 1
End If
Loop
i% = i% + 1
divide_item_mark0:
Loop
If i1$ = "" And i2$ = "" Then
 divide_item = "1"
ElseIf i2$ = "" Then
divide_item = i1$
ElseIf i1$ = "" Then
divide_item = "(1/" + i2$ + ")"
Else
 divide_item = "(" + i1$ + "/" + i2$ + ")"
End If
ElseIf n1% > 0 And n2 = 0 Then
divide_item = divide_item(Mid$(i1$, 1, n1% - 1), _
      time_item(i2$, Mid$(i1$, n1% + 1, Len(i1$) - n1%)), "", "")
ElseIf n1% = 0 And n2 > 0 Then
divide_item = divide_item(time_item(i1$, _
               Mid$(i2$, n2% - 1, Len(i2$) - n2%)), _
                  Mid$(i2$, 1, n2% - 1), "", "")
Else
divide_item = divide_item(time_item(Mid$(i1$, 1, n1% - 1), _
      Mid$(i2$, n2% - 1, Len(i2$) - n2%)), _
         time(Mid$(i1$, n1% + 1, Len(i1$) - n1%), Mid$(i2$, 1, n2% - 1)), "", "")
End If
Exit Function
divide_item_error:
divide_item = "F"
End Function

Private Function is_item(s As String, para As String, item0 As String) As Boolean
Dim p%, l%
If InStr(1, s, "(", 0) = 0 And InStr(1, s, "+", 0) = 0 And _
 InStr(1, s, "-", 0) = 0 And InStr(1, s, "/", 0) = 0 Then
  is_item = True
   p% = InStr(1, s, "*", 0)
    l% = Len(s)
    If p% > 0 Then
      If p% = 1 Then
      para = "1"
      Else
      para = Mid$(s, 1, p% - 1)
      End If
      item0 = Mid$(s, p% + 1, l - p%)
    Else
     If Mid$(s, 1, 1) > "A" Then
     para = "1"
      item0 = s
     Else
     para = s
     item0 = ""
     End If
    End If
End If
End Function
Private Sub read_para_from_item(ByVal s As String, _
      pA As String, it As String)
Dim i%, i_n%
Dim ch As String
i% = InStr(1, s, "*", 0) '乘号
If i% > 0 Then ' 有乘号
 pA = Mid$(s, 1, i% - 1)
  If Mid$(pA, 1, 1) = "-" Then
   pA = "@" + Mid$(pA, 2, Len(pA) - 1)
  End If
  it = Mid$(s, i% + 1, Len(s) - i%)
Else ' 无乘号
 For i% = 1 To Len(s)
  ch = Mid(s, i%, 1)
   If ch <> "|" And (ch = "[" Or ch >= "A" Or ch = "\") Then '系数分界
    i_n% = i%
     GoTo read_para_from_item_mark0
   End If
 Next i%
read_para_from_item_mark0:
 If i_n% = 0 Then '纯数
  pA = s
  it = "1"
 ElseIf i_n% = 1 Then
  pA = "1"
  it = s
 ElseIf i_n% >= 2 Then
  If Mid$(s, i_n% - 1, 1) = "'" Then
  pA = Mid(s, 1, i_n% - 2)
  it = Mid(s, i_n% - 1, Len(s) - i_n% + 2)
  Else
  pA = Mid(s, 1, i_n% - 1)
  it = Mid(s, i_n%, Len(s) - i_n% + 1)
  End If
   If pA = "" Or pA = "+1" Or pA = "#1" Or pA = "+" Or pA = "#" Then
      pA = "1"
   ElseIf pA = "-" Or pA = "@" Or pA$ = "-1" Then
    pA = "@1"
   End If
 'Else
 ' pa = Mid(S, 1, i_n% - 1)
 ' it = Mid(S, i_n%, Len(S) - i_n% + 1)
 End If
 End If
End Sub

Public Function time_string(ByVal s1 As String, ByVal S2 As String, is_simple As Boolean, _
          cal_float As Boolean) As String 's1和s2相乘
Dim t(1) As Integer
Dim it(3) As String
Dim i%, j%
Dim f(1) As String
Dim fs(2) As String
Dim s(5) As String
Dim ts As String
Dim ch As String * 1
Dim v As Variant
Dim need_reduce As Boolean
On Error GoTo time_string_error
If s1 = "" Or S2 = "" Then
 time_string = "F"
  Exit Function
ElseIf InStr(1, s1, ".", 0) > 0 Or InStr(1, S2, ".", 0) > 0 Then
   cal_float = True
End If
If InStr(1, s1, "F", 0) > 0 Or InStr(1, S2, "F", 0) > 0 Or s1 = "" Or S2 = "" Then
time_string_error:
 time_string = "F"
  Exit Function
ElseIf s1 = "1" Then
 If Mid$(S2, 1, 1) = "@" Then
  time_string = "-" & Mid$(S2, 2, Len(S2) - 1)
 ElseIf Mid$(S2, 1, 1) = "#" Then
  time_string = Mid$(S2, 2, Len(S2) - 1)
 Else
  time_string = S2
 End If
   Exit Function
ElseIf s1 = "0" Then
 time_string = "0"
   Exit Function
ElseIf S2 = "1" Then
 If Mid$(s1, 1, 1) = "@" Then
 time_string = "-" & Mid$(s1, 2, Len(s1) - 1)
 ElseIf Mid$(s1, 1, 1) = "#" Then
 time_string = Mid$(s1, 2, Len(s1) - 1)
 Else
 time_string = s1
 End If
   Exit Function
ElseIf S2 = "" Or S2 = "0" Then
 time_string = "0"
   Exit Function
ElseIf s1 = "P" Then
 If S2 = "P" Then
  time_string = "P"
   Exit Function
 ElseIf S2 = "N" Then
  time_string = "N"
   Exit Function
 Else
  If val(value_string(S2)) > 0 Then
   time_string = "P"
    Exit Function
  ElseIf val(value_string(S2)) < 0 Then
   time_string = "N"
    Exit Function
  Else
   time_string = "F"
    Exit Function
  End If
 End If
ElseIf s1 = "N" Then
 If S2 = "P" Then
  time_string = "N"
   Exit Function
 ElseIf S2 = "N" Then
  time_string = "P"
   Exit Function
 Else
  If val(value_string(S2)) > 0 Then
   time_string = "N"
    Exit Function
  ElseIf val(value_string(S2)) < 0 Then
   time_string = "P"
    Exit Function
  Else
   time_string = "F"
    Exit Function
  End If
 End If
ElseIf S2 = "P" Or S2 = "N" Then
 time_string = time_string(S2, s1, True, False)
  Exit Function
End If
If regist_data.run_type = 1 Then
   If (InStr(1, s1, "U", 0) > 0 Or InStr(1, s1, "V", 0) > 0) And _
        (InStr(1, S2, "U", 0) > 0 Or InStr(1, S2, "V", 0) > 0) Then
         need_reduce = True
   End If
End If
t(0) = string_type(s1, fs(0), s(0), s(1), s(2))
 If s1 = S2 Then
  If t(0) = 0 Then
   If s(2) = "" Then
    time_string = time_string( _
      time_para(s(0), s(0), False, cal_float), time_item(s(1), s(1)), is_simple, cal_float)
   Else
    time_string = add_string(time_string(s1, fs(0), False, cal_float), _
         time_string(s1, s(2), False, cal_float), is_simple, cal_float)
   End If
  ElseIf t(0) = 3 Then
    time_string = divide_string( _
       time_string(s(0), s(0), False, cal_float), time_string(s(1), s(1), False, cal_float), _
         is_simple, cal_float)
  End If
  Exit Function
 End If
 t(1) = string_type(S2, fs(1), s(3), s(4), s(5))
If t(0) = 3 And t(1) = 3 Then
  If gcd_for_string(s(3), s(1), "", f(0), f(1), True) Then
   s(3) = f(0)
   s(1) = f(1)
  End If
  If gcd_for_string(s(0), s(4), "", f(0), f(1), True) Then
   s(0) = f(0)
   s(4) = f(1)
  End If
   time_string = divide_string(time_string(s(0), s(3), False, cal_float), _
         time_string(s(1), s(4), False, cal_float), is_simple, cal_float)
          Exit Function
ElseIf t(0) = 3 And t(1) = 0 Then
 If gcd_for_string(S2, s(1), "", f(0), f(1), True) Then
  S2 = f(0)
  s(1) = f(1)
 End If
 time_string = divide_string(time_string(S2, s(0), False, cal_float), _
         s(1), is_simple, cal_float)
          Exit Function
ElseIf t(1) = 3 And t(0) = 0 Then
  time_string = time_string(S2, s1, is_simple, cal_float)
ElseIf t(0) = 0 And t(1) = 0 Then
  If s(2) = "" And s(5) = "" Then
    If InStr(1, s(0), ".") > 0 Or InStr(1, s(3), ".") > 0 Then
     cal_float = True
    End If
    s(0) = time_para(s(0), s(3), is_simple, cal_float)
      If s(1) = "1" Or s(4) = "1" Then
        s(1) = time_item(s(1), s(4))
       time_string = combine_item_with_para(s(0), s(1), is_simple)
      Else
        s(1) = time_item(s(1), s(4))
       time_string = time_string(s(0), s(1), is_simple, cal_float)
      End If
  ElseIf s(2) = "" Then
    time_string = add_string(time_string(s1, fs(1), False, cal_float), _
        time_string(s1, s(5), False, cal_float), False, cal_float)
  ElseIf s(2) <> "" Then
    time_string = add_string(time_string(fs(0), S2, False, cal_float), _
        time_string(s(2), S2, False, cal_float), is_simple, cal_float)
  End If
 End If
 If regist_data.run_type = 1 And need_reduce Then
  If InStr(1, time_string, "UU", 0) > 0 Or InStr(1, time_string, "UV", 0) > 0 Or _
       InStr(1, time_string, "VV", 0) > 0 Then
    Call reduce_string_by_string(time_string)
  End If
 End If
 If is_simple = True Then
 time_string = simple_string_(time_string, "", "", cal_float)
 End If
End Function


Public Function add_string(ByVal s1 As String, ByVal S2 As String, is_simple As Boolean, _
            cal_float As Boolean) As String 's1和s2 相加
Dim i%, j%, k%
Dim s(6) As String
Dim t(1) As Integer
Dim ts As String
Dim S_p As para_item_type
On Error GoTo add_string_error
If s1 = "" Or S2 = "" Then
 add_string = "F"
  Exit Function
ElseIf InStr(1, s1, ".", 0) > 0 Or InStr(1, S2, ".", 0) > 0 Then
   cal_float = True
End If
If InStr(1, s1, "F", 0) Or InStr(1, S2, "F", 0) > 0 Then
 add_string = "F"
  Exit Function
ElseIf s1 = "0" Then
 If Mid$(S2, 1, 1) = "@" Then
  add_string = "-" & Mid$(S2, 2, Len(S2) - 1)
 ElseIf Mid$(S2, 1, 1) = "#" Then
  add_string = Mid$(S2, 2, Len(S2) - 1)
 Else
  add_string = S2
 End If
  Exit Function
ElseIf S2 = "0" Then
  If Mid$(s1, 1, 1) = "@" Then
  add_string = "-" & Mid$(s1, 2, Len(s1) - 1)
 ElseIf Mid$(s1, 1, 1) = "#" Then
  add_string = Mid$(s1, 2, Len(s1) - 1)
 Else
  add_string = s1
 End If
 Exit Function
ElseIf s1 = "P" Then
 If S2 = "P" Then
  add_string = "P"
   Exit Function
 End If
ElseIf s1 = "N" Then
 If S2 = "N" Then
  add_string = "N"
   Exit Function
 End If
ElseIf s1 = S2 Then
 add_string = time_string(s1, "2", is_simple, cal_float)
  Exit Function
End If
 t(0) = string_type(s1, "", s(0), s(1), s(2))
  t(1) = string_type(S2, ts, s(3), s(4), s(5))
If t(0) = 3 And t(1) = 3 Then
 If s(1) = s(4) Then
  add_string = divide_string(add_string(s(0), s(3), False, cal_float), _
       s(1), is_simple, cal_float)
 Else
 If lcd_for_string(s(1), s(4), s(6), s(2), s(5), False) = False Then
    add_string = time_string(s(0), s(4), False, cal_float)
    add_string = add_string(add_string, time_string(s(1), s(3), False, cal_float), False, cal_float)
    add_string = divide_string(add_string, time_string(s(4), s(1), False, cal_float), True, cal_float)
   Exit Function
 End If
  add_string = divide_string(add_string(time_string(s(0), s(5), False, cal_float), _
      time_string(s(3), s(2), False, cal_float), False, cal_float), s(6), is_simple, cal_float)
 End If
ElseIf t(0) = 3 Then '分式
 add_string = divide_string(add_string(time_string(s(1), S2, False, cal_float), _
        s(0), False, cal_float), s(1), is_simple, cal_float)
ElseIf t(1) = 3 Then '分式
 add_string = divide_string(add_string(time_string(s(4), s1, False, cal_float), _
       s(3), False, cal_float), s(4), is_simple, cal_float)
 'add_string = add_string(S2, s1)
ElseIf t(0) = 0 And t(1) = 0 Then '整式
 If InStr(1, s1, ".") > 0 Or InStr(1, S2, ".") > 0 Then
  cal_float = True
  s1 = value_string(s1)
   S2 = value_string(S2)
 t(0) = string_type(s1, "", s(0), s(1), s(2))
  t(1) = string_type(S2, "", s(3), s(4), s(5))
 End If
 If s(5) = "" Then
 If read_string(s1, S_p) = False Then
     GoTo add_string_error
 End If
i% = 0
Do While i% < S_p.last_it
   ' For i% = 0 To s_p.last_it - 1
 If s(4) < S_p.it(i%) Then 'Or s_p.it(i%) = "1") Then '按项排序
  ReDim Preserve S_p.pA(S_p.last_it) As String
   ReDim Preserve S_p.it(S_p.last_it) As String
        For j% = S_p.last_it - 1 To i% Step -1
         S_p.pA(j% + 1) = S_p.pA(j%)
          S_p.it(j% + 1) = S_p.it(j%)
        Next j%
         S_p.pA(i%) = s(3)
          S_p.it(i%) = s(4)
          S_p.last_it = S_p.last_it + 1
            GoTo add_string_mark1
  ElseIf s(4) = S_p.it(i%) Then
        S_p.pA(i%) = add_para(S_p.pA(i%), s(3), is_simple, cal_float)
         If S_p.pA(i%) = "0" Then
          For k% = i% To S_p.last_it - 2
           S_p.pA(k%) = S_p.pA(k% + 1)
            S_p.it(k%) = S_p.it(k% + 1)
          Next k%
         S_p.last_it = S_p.last_it - 1
   End If
            GoTo add_string_mark1
       End If
     i% = i% + 1
     Loop
  ReDim Preserve S_p.pA(S_p.last_it) As String
  ReDim Preserve S_p.it(S_p.last_it) As String
  S_p.last_it = S_p.last_it + 1
          S_p.pA(S_p.last_it - 1) = s(3)
          S_p.it(S_p.last_it - 1) = s(4)
          ' s_p.last_it = s_p.last_it + 1
add_string_mark1:
    add_string = string_from_para_item(S_p)
Else
add_string = add_string(add_string(s1$, ts, False, cal_float), s(5), is_simple, cal_float)
End If
End If
If is_simple = True Then
add_string = simple_string_(add_string, "", "", cal_float)
End If
Exit Function
add_string_error:
add_string = "F"
End Function

Public Function minus_string(ByVal s1 As String, ByVal S2 As String, _
        is_simple As Boolean, cal_float As Boolean) As String 's1和s2 相减
If s1 = S2 Then
 minus_string = "0"
Else
 minus_string = add_string(s1, time_string("-1", S2, False, cal_float), is_simple, cal_float)
End If
End Function
Private Function add_para(p1 As String, p2 As String, is_simple As Boolean, _
           cal_float As Boolean) As String
Dim i%, j%, k%
Dim t(1) As Integer
 Dim ts(1) As String
  Dim s(5) As String
   Dim t_s$
Dim it1 As para_item_type
Dim it2 As para_item_type
On Error GoTo add_para_error
If InStr(1, p1, "F", 0) > 0 Or InStr(1, p1, "F", 0) > 0 Then
 add_para = "F"
  Exit Function
ElseIf p1 = "" Or p1 = "0" Then
 If Mid$(p2, 1, 1) = "@" Then
 add_para = "-" & Mid$(p2, 2, Len(p2) - 1)
 ElseIf Mid$(p2, 1, 1) = "#" Then
 add_para = Mid$(p2, 2, Len(p2) - 1)
 Else
 add_para = p2
 End If
Exit Function
ElseIf p2 = "" Or p2 = "0" Then
 If Mid$(p1, 1, 1) = "@" Then
 add_para = "-" & Mid$(p1, 2, Len(p1) - 1)
 ElseIf Mid$(p1, 1, 1) = "#" Then
 add_para = Mid$(p1, 2, Len(p1) - 1)
 Else
 add_para = p1
 End If
 Exit Function
ElseIf p1 = p2 Then
 add_para = time_string("2", p1, True, True)
End If
t(0) = para_type(p1, ts(0), s(0), s(1), s(4))
 t(1) = para_type(p2, ts(1), s(2), s(3), s(5))
If t(0) = 2 Or t(1) = 2 Then
 '****************************
  If cal_float = True Then
   add_para = str_(val_(value_para(p1)) + val_(value_para(p2)))
  Else
   add_para = "F"
  End If
ElseIf t(1) = 3 Then
 '整数分数相加
  add_para = divide_para(add_para(time_para(p1, s(3), False, cal_float), _
       s(2), False, cal_float), s(3), True, cal_float)
ElseIf t(0) = 3 Then
  add_para = add_para(p2, p1, True, cal_float)
ElseIf (t(0) = 0 Or t(0) = 1) And (t(1) = 0 Or t(1) = 1) Then
 If s(4) = "" Then
    Call read_para(p2, it1)
 Do While i% < it1.last_it
  If s(1) > it1.it(i%) Then
   ReDim Preserve it1.pA(it1.last_it) As String
    ReDim Preserve it1.it(it1.last_it) As String
    For j% = it1.last_it To i% + 1 Step -1
     it1.pA(j%) = it1.pA(j% - 1)
     it1.it(j%) = it1.it(j% - 1)
    Next j%
     it1.last_it = it1.last_it + 1
     it1.pA(i%) = s(0)
     it1.it(i%) = s(1)
    GoTo add_para_loop_out
   ElseIf s(1) = it1.it(i%) Then
   it1.pA(i%) = str_(val_(s(0)) + val_(it1.pA(i%)))
   If it1.pA(i%) = "0" Then
     For j% = i% To it1.last_it - 2
     it1.pA(j%) = it1.pA(j% + 1)
     it1.it(j%) = it1.it(j% + 1)
    Next j%
     it1.last_it = it1.last_it - 1
  End If
    GoTo add_para_loop_out
  Else
  i% = i% + 1
  End If
 Loop
    ReDim Preserve it1.pA(it1.last_it) As String
    ReDim Preserve it1.it(it1.last_it) As String
     it1.pA(it1.last_it) = s(0)
     it1.it(it1.last_it) = s(1)
     it1.last_it = it1.last_it + 1
add_para_loop_out:
add_para = para_from_para_item(it1)
 ElseIf s(4) <> "" Then
   add_para = add_para(s(4), _
      add_para(ts(0), p2, False, cal_float), False, cal_float)
 End If
 '**************************************************************
End If
If is_simple = True Then
add_para = simple_para(add_para, "", "")
End If
Exit Function
add_para_error:
add_para = "F"
End Function

Private Function no_brace(s$, n1%, n2%) As Boolean
'判断n1%,n2% 在括号外
Dim i%, k%
Dim c$
k% = 0
For i% = n1% To n2%
c$ = Mid$(s$, i%, 1)
If c$ = "(" Then
k% = k% + 1
ElseIf c$ = ")" Then
k% = k% - 1
End If
Next i%
If k% = 0 Then
no_brace = True
End If
End Function

Public Function is_brace(s As String, n1%, n2%)
'判断n1%,n2%是否是()
Dim i%, k%
If Mid$(s, n1%, 1) = "(" And Mid$(s, n2%, 1) = ")" Then
For i% = n1% + 1 To n2% - 1
If Mid$(s, i%, 1) = "(" Then
 k% = k% + 1
 ElseIf Mid$(s, i%, 1) = ")" Then
 k% = k% - 1
  If k% < 0 And i% < n2% Then
  is_brace = False
   Exit Function
  End If
 End If
Next i%
If k% = 0 Then
 is_brace = True
End If
End If
End Function
Private Function para_from_para_item(p_i As para_item_type) As String
Dim i%
If p_i.last_it = 0 Then
para_from_para_item = "0"
Else
For i% = 0 To p_i.last_it - 1
para_from_para_item = combine_para_or_string_for_add( _
 para_from_para_item, combine_para_for_item(p_i.pA(i%), _
   p_i.it(i%)), "para")
Next i%
End If
End Function
Public Function para_type(ByVal p As String, p0 As String, _
    s1 As String, S2 As String, s3 As String) As Integer
'0整式 #
'1根式 ''
'2浮点数'.
'3 分式'/,&
Dim m%, k%, m_%, k_%
Dim t_p As String
Dim ts$
Dim tn(3) As Long
If p = "" Or p = "0" Then '空串
 para_type = 0
  s1 = "0"
   S2 = "0"
    s3 = ""
  Exit Function
End If
p = unsimple_para(p) '去括号
m% = InStr(1, p, "&", 0) '除号
If m% = 0 Then
 m% = InStr(1, p, "/", 0) '除号
End If
k% = InStr(1, p, ".", 0) '浮点
If m% > 0 Then '有除号
 para_type = 3
  s1 = Mid$(p, 1, m% - 1)
   S2 = Mid$(p, m% + 1, Len(p) - m%)
    s3 = ""
ElseIf k% > 0 And k% <> Len(p) Then '有浮点
 s1 = p
  S2 = "1"
   s3 = ""
 para_type = 2
Else
   m% = InStr(2, p, "#", 0) '
    m_% = InStr(2, p, "+", 0)
If (m_% > 0 And m_% < m%) Or m% = 0 Then
    m% = m_%
End If
    k% = InStr(2, p, "@", 0)
     k_% = InStr(2, p, "-", 0)
If (k_% > 0 And k_% < k%) Or k% = 0 Then
    k% = k_%
End If
If (k% > 0 And k% < m%) Or m% = 0 Then
      m% = k%
End If
If m% > 0 Then '第一个加减
  t_p = Mid$(p, 1, m% - 1)
   p0 = t_p
    s3 = Mid$(p, m%, Len(p) - m% + 1)
     If Mid$(s3, 1, 1) = "#" Then
      s3 = Mid$(s3, 2, Len(s3) - 1)
     End If
Else
   t_p = p
    p0 = t_p
     s3 = ""
End If
    m% = InStr(1, t_p, "'", 0) '根号
If m% > 0 Then
 Call read_sqr_string_from_string(t_p, m%, m%, S2$, "", s1$)
   If s1 = "" Or s1 = "#" Or s1 = "+" Then
    s1 = "1"
   ElseIf s1 = "@" Or s1 = "-" Then
    s1 = "@1"
   End If
   If ts$ <> "" Then
    If s1 = "1" Then
     s1$ = ts$
    ElseIf s1 = "@" Then
     s1$ = "@" & ts$
    Else
     If ts$ <> "1" Then
      s1$ = ts$ & s1$
     End If
    End If
   End If
   If s3 = "" Then
    para_type = 1
   Else
    para_type = 0
   End If
   tn(0) = val(S2$)
    Call sq_root(tn(0), tn(1), tn(2))
     If tn(1) <> "1" Then
        S2$ = Trim(str(tn(2)))
         s1$ = Trim(str(val(s1$) * tn(1)))
     End If
Else
  s1 = t_p
   S2 = "1"
    para_type = 0
End If
   End If
End Function

Private Function minus_para(p1 As String, p2 As String, is_simple As Boolean, _
          cal_float As Boolean) As String
If p1 = p2 Then
 minus_para = "0"
Else
 minus_para = add_para(p1, time_para("@1", p2, False, cal_float), is_simple, cal_float)
End If
End Function
Public Function add_sign_no(start%, s$, t_y As String) As Integer
' 第一个加减号
Dim m%, n%, k%, st%
Dim ts$
Dim ty As Boolean
st% = start%
k% = InStr(st%, s$, ")", 0)
Do While ty = False
If t_y = "string" Then
m% = InStr(st%, s$, "+", 0)
 n% = InStr(st%, s$, "-")
Else
m% = InStr(st%, s$, "#", 0)
 n% = InStr(st%, s$, "@", 0)
End If
If n% = 0 And m% = 0 Then
If k% > 0 Then
add_sign_no = k% - 1
Else
add_sign_no = Len(s$) + 1
End If
 Exit Function
ElseIf n% = 0 Then
add_sign_no = m%
ElseIf m% = 0 Then
add_sign_no = n%
Else
add_sign_no = min(m%, n%)
End If
ty = no_brace(s$, start%, add_sign_no)
Loop
End Function

Private Function time_para(ByVal p1 As String, ByVal p2 As String, is_simple As Boolean, _
         cal_float As Boolean) As String
Dim t(1) As Integer
Dim tp(1) As String
Dim s(6) As String
Dim fs(2) As String
Dim tn(1) As Long
On Error GoTo time_para_error
If InStr(1, p1, "F", 0) > 0 Or InStr(1, p2, "F", 0) > 0 Then
 time_para = "F"
  Exit Function
End If
If p1 = "1" Then
 If Mid$(p2, 1, 1) = "@" Then
 time_para = "-" & Mid$(p2, 2, Len(p2) - 1)
 ElseIf Mid$(p2, 1, 1) = "#" Then
 time_para = Mid$(p2, 2, Len(p2) - 1)
 Else
 time_para = p2
 End If
  time_para = simple_para(time_para, "", "")
  Exit Function
ElseIf p2 = "1" Then
  If Mid$(p1, 1, 1) = "@" Then
 time_para = "-" & Mid$(p1, 2, Len(p1) - 1)
 ElseIf Mid$(p1, 1, 1) = "#" Then
 time_para = Mid$(p1, 2, Len(p1) - 1)
 Else
 time_para = p1
 End If
 time_para = simple_para(time_para, "", "")
   Exit Function
ElseIf p1 = "0" Or p2 = "0" Or p1 = "" Or p2 = "" Then
 time_para = "0"
  Exit Function
End If
t(0) = para_type(p1, tp(0), s(0), s(1), s(2))
 t(1) = para_type(p2, tp(1), s(3), s(4), s(5))
  If t(0) = 2 Or t(1) = 2 Then
   If cal_float = True Then
    time_para = str_(val_(value_para(p1)) * val_(value_para(p2)))
   Else
    time_para = "F"
   End If
  ElseIf t(0) = 3 Then
    time_para = divide_para(time_para(p2, s(0), False, cal_float), s(1), False, cal_float)
  ElseIf t(1) = 3 Then
   time_para = time_para(p2, p1, False, cal_float)
  ElseIf (t(0) = 0 Or t(0) = 1) And _
            (t(1) = 0 Or t(1) = 1) Then
   If s(2) = "" And s(5) = "" Then
    If s(1) = s(4) Then '相同根式
     s(3) = time_para(s(4), s(3), is_simple, cal_float)
     s(4) = "1"
     s(1) = "1"
    Else
     Call sqr_para(str_(val_(s(1)) * val_(s(4))), s(1), s(4), cal_float)
    End If
        time_para = combine_para_for_item(str_(val_(s(0)) * val_(s(3)) * _
          val_(s(1))), s(4))
   ElseIf s(2) <> "" Then
    time_para = add_para(time_para(tp(0), p2, False, cal_float), _
        time_para(s(2), p2, False, cal_float), False, cal_float)
   ElseIf s(2) = "" Then
    time_para = add_para(time_para(p1, tp(1), False, cal_float), _
        time_para(p1, s(5), False, cal_float), False, cal_float)
   End If
  End If
If is_simple = True Then
time_para = simple_para(time_para, "", "")
End If
Exit Function
time_para_error:
time_para = "F"
End Function

Public Function is_single_para(s As String, s1 As String, S2 As String) As Boolean
'整数乘根号
Dim p(1) As Integer
If InStr(1, s, "#", 0) = 0 And InStr(1, s, "$", 0) = 0 And InStr(1, s, "&", 0) = 0 Then
 is_single_para = True
  p(0) = InStr(1, s, "%", 0)
   p(1) = InStr(1, s, "'", 0)
    If p(0) = 0 And p(1) = 0 Then
      s1 = s
    ElseIf p(0) = 0 Then
     S2 = Mid$(s, 2, Len(s) - 1)
    Else
     s1 = Mid$(s, 1, p(0) - 1)
      S2 = Mid$(s, p(1) + 1, Len(s) - p(1))
    End If
End If
End Function

Private Function divide_para(ByVal p1 As String, ByVal p2 As String, is_simple As Boolean, _
          cal_float As Boolean) As String
Dim t(1) As Integer
Dim tp(1) As String
Dim tn(7) As Long
Dim ts(7) As String
Dim fs(2) As String
Dim s(5) As String
Dim finish As Boolean
Dim i%, j%
On Error GoTo divide_para_error
If p1 = p2 Then
divide_para = "1"
divide_para = simple_para(divide_para, "", "")
Exit Function
ElseIf p1 = "0" Then
divide_para = "0"
divide_para = simple_para(divide_para, "", "")
Exit Function
ElseIf p1 = "F" Or p1 = "" Or p2 = "0" Or p2 = "" Or p2 = "F" Then
divide_para = "F"
 Exit Function
ElseIf p2 = "1" Then
 divide_para = p1
divide_para = simple_para(divide_para, "", "")
  Exit Function
ElseIf p2 = "-1" Or p2 = "@1" Then
 divide_para = time_para("-1", p1, is_simple, cal_float)
      Exit Function
End If
  t(0) = para_type(p1, tp(0), s(0), s(1), "")
    t(1) = para_type(p2, tp(1), s(2), s(3), "")
 If t(1) = 2 Or t(0) = 2 Then
   divide_para = str_(val_(value_para(p1)) / val_(value_para(p2)))
 ElseIf t(1) = 1 Then
  Call rational_para(p2, fs(0), fs(1))
  divide_para = divide_para(time_para(p1, fs(0), False, cal_float), fs(1), False, cal_float)
 ElseIf t(0) = 3 And t(1) = 3 Then
 divide_para = divide_para(time_string(s(3), s(0), False, cal_float), _
                time_para(s(1), s(2), False, cal_float), False, cal_float)
 ElseIf t(0) = 3 Then
 divide_para = divide_para(s(0), time_para(s(1), p2, False, cal_float), False, cal_float)
 ElseIf t(1) = 3 Then
 divide_para = divide_para(time_string(s(3), p1, False, cal_float), s(2), False, cal_float)
 ElseIf t(0) = 1 Then
  If t(1) = 0 Then
   If s(3) = "1" Then
    tn(0) = val_(s(0))
     tn(1) = val_(p2)
   Call simple_multi_long(tn(1), tn(0), 0, 0, 0, 0, 0, 0, 0, 2, 0)
    ts(0) = str_(tn(0))
     ts(1) = str_(tn(1))
      If ts(0) = "" Or ts(0) = "1" Then
       ts(0) = "'" + s(1)
      ElseIf ts(0) = "-1" Or ts(0) = "@-" Then
       ts(0) = "-'" + s(1)
      ElseIf s(1) <> "1" And s(1) <> "" Then
       ts(0) = ts(0) + "'" + s(1)
      End If
      If ts(1) = "" Or ts(1) = "1" Then
       divide_para = ts(0)
      Else
      divide_para = ts(0) + "&" + ts(1)
      End If
  Else
   If para_type(s(3), "", s(4), s(5), "") = 1 Then
    Call rational_para(p2, s(2), s(3))
     divide_para = divide_para(time_para(p1, s(2), False, cal_float), s(3), False, cal_float)
  Else
    divide_para = str_(val_(value_para(p1)) / _
                    val_(value_para(p2)))
   End If
  End If
  End If
 ElseIf t(0) = 0 Then
  If t(1) = 0 Then
   If s(3) = "1" Then '分母是整数
    If s(1) = "1" Then '分子是整数
     tn(0) = val_(p1)
      tn(1) = val_(p2)
       Call simple_two_long(tn(1), tn(0), 0)
        If tn(1) = "1" Then
         divide_para = str_(tn(0))
        Else
          divide_para = str_(tn(0)) + "&" + str_(tn(1))
        End If
    Else
     Call simple_para(p1, "", "") '提取因子
      i% = InStr(1, p1, "(", 0)
        If i% > 0 Then
          ts(0) = Mid$(p1, 1, i% - 1)
          ts(1) = Mid$(p1, i%, Len(p1) - i% + 1)
           If ts(0) = "" Then
            ts(0) = "1"
           ElseIf ts(0) = "-" Then
            ts(0) = "-1"
           End If
            tn(1) = val_(p2)
            tn(0) = val_(ts(0))
        If tn(0) <> 0 Then
          Call simple_two_long(tn(0), tn(1), 0)
          If tn(0) = "1" Then
           divide_para = ts(1)
          ElseIf tn(0) = "-1" Then
           divide_para = "-" + ts(1)
          Else
           divide_para = str_(tn(0)) + ts(1)
          End If
          If tn(1) <> "1" Then
           divide_para = divide_para + _
            "&" + str_(tn(1))
          End If
        Else
         divide_para = p1 + "&" + p2
          Exit Function
        End If
      Else
       If InStr(1, p1, "@", 0) = 0 And InStr(1, p1, "#", 0) = 0 Then
           divide_para = p1 + "&" + p2
       Else
           divide_para = "(" + p1 + ")" + "&" + p2
       End If
      End If
End If
Else
       Call rational_para(p2, ts(0), ts(1))
          divide_para = divide_para(time_para(p1, ts(0), False, cal_float), ts(1), _
             False, cal_float)
End If
End If
End If
If is_simple = True Then
divide_para = simple_para(divide_para, "", "")
End If
Exit Function
divide_para_error:
divide_para = "F"
End Function



Public Function determinant(ByVal A11$, ByVal A12$, _
    ByVal A21$, ByVal A22$) As String
determinant = minus_string(time_string(A11$, A22$, False, False), _
      time_string(A12$, A21$, False, False), True, False)
End Function
Public Function divide_string(ByVal s1 As String, ByVal S2 As String, is_simple As Boolean, _
                  cal_float As Boolean) As String 's1和s2 相除
Dim te_s(1) As String
Dim i%, j%, tn%, sqr_no%
Dim fs(1) As String
Dim fc(1) As String
Dim s(6) As String
Dim t(1) As Integer
Dim ts(8) As String
Dim S_p_i As para_item_type
Dim S_p_i1 As para_item_type
Dim pA(1) As String
Dim f(5) As String
Dim f1(3) As String
Dim f2(3) As String
Dim is_t(1) As Boolean
Dim sq As String
Dim v_string_type As Byte
'*******
'Type para_item_type
 ' it(8) As String
  ' pa(8) As String
   ' last_it As Integer
 'End Type
On Error GoTo divide_string_error
If InStr(1, s1, ".", 0) > 0 Or InStr(1, S2, ".", 0) > 0 Then
   cal_float = True
End If
'*****************
If (InStr(1, s1, "UU", 0) > 0 Or InStr(1, s1, "VV", 0) > 0 Or InStr(1, s1, "UV", 0) > 0) And _
     (InStr(1, S2, "UU", 0) > 0 Or InStr(1, S2, "VV", 0) > 0 Or InStr(1, S2, "UV", 0) > 0) Then
    v_string_type = 1
ElseIf (InStr(1, s1, "UU", 0) > 0 Or InStr(1, s1, "VV", 0) > 0 Or InStr(1, s1, "UV", 0) > 0) And _
     (InStr(1, S2, "U", 0) > 0 Or InStr(1, S2, "V", 0) > 0) Then
    v_string_type = 2
ElseIf (InStr(1, s1, "U", 0) > 0 Or InStr(1, s1, "V", 0) > 0) And _
     (InStr(1, S2, "UU", 0) > 0 Or InStr(1, S2, "VV", 0) > 0 Or InStr(1, S2, "UV", 0) > 0) Then
    v_string_type = 3
ElseIf (InStr(1, s1, "U", 0) > 0 Or InStr(1, s1, "V", 0) > 0) And _
     (InStr(1, S2, "U", 0) > 0 Or InStr(1, S2, "V", 0) > 0) Then
    v_string_type = 4
End If
If InStr(1, s1, "F", 0) > 0 Or InStr(1, S2, "F", 0) > 0 Or S2 = "0" Or s1 = "" Or S2 = "" Then
 divide_string = "F"
  Exit Function
ElseIf s1 = "0" Then
 divide_string = "0"
  Exit Function
ElseIf S2 = "1" Then
  divide_string = s1
   Exit Function
ElseIf S2 = "-1" Or S2 = "@1" Then
  divide_string = time_string(s1, "-1", is_simple, cal_float)
   Exit Function
ElseIf s1 = S2 Then
 divide_string = "1"
  Exit Function
ElseIf S2 = "P" Then
 If s1 = "P" Then
  divide_string = "P"
   Exit Function
 ElseIf s1 = "N" Then
  divide_string = "N"
   Exit Function
 Else
  If val(value_string(s1)) > 0 Then
   divide_string = "P"
    Exit Function
  ElseIf val(value_string(s1)) < 0 Then
   divide_string = "N"
    Exit Function
  Else
   divide_string = "F"
    Exit Function
  End If
 End If
ElseIf S2 = "N" Then
 If s1 = "P" Then
  divide_string = "N"
   Exit Function
 ElseIf s1 = "N" Then
  divide_string = "P"
   Exit Function
 Else
  If val(value_string(s1)) > 0 Then
   divide_string = "N"
    Exit Function
  ElseIf val(value_string(s1)) < 0 Then
   divide_string = "P"
    Exit Function
  Else
   divide_string = "F"
    Exit Function
  End If
 End If
End If
'分离两个树
If InStr(1, s1, ".", 0) > 0 Then
 cal_float = True
' s1 = value_string(s1)
End If
If InStr(1, S2, ".", 0) > 0 Then
 cal_float = True
'  S2 = value_string(S2)
End If
'If InStr(1, s1, ".", 0) > 0 Or InStr(1, s1, ".", 0) > 0 Then
 'cal_float = True
'End If
t(0) = string_type(s1, fs(0), s(0), s(1), s(2))
 t(1) = string_type(S2, fs(1), s(3), s(4), s(5))
  sqr_no% = InStr(1, s(4), "[", 0)
If t(0) = 0 And t(1) = 0 Then
If s(5) = "" Then
 If (sqr_no% > 0 Or _
      (InStr(1, s(4), "+", 0) > 1 And InStr(1, s(4), "-", 0) > 1)) Then
      '除数是根式
   Call do_factor1(s1, f1(0), f1(1), f1(2), f1(3), 0) '因式分解被除数
   If sqr_no% > 0 Then
    sq = read_sqr_from_string(s(4), 0, s(0))
     'tn% = from_char_to_no(s(4))
    For i% = 0 To 3
      If f1(i%) = sq Then 'squre_root_string(tn%) Then
       f1(i%) = s(4)
       s(4) = "1"
      ElseIf f1(i%) = S2 Then
       f1(i%) = "1"
       s(4) = "1"
      End If
    Next i%
     divide_string = time_string(f1(0), f1(1), False, cal_float)
     divide_string = time_string(divide_string, f1(2), False, cal_float)
     divide_string = time_string(divide_string, f1(3), is_simple, cal_float)
     If s(4) = "1" Then
        divide_string = divide_string(divide_string, _
         s(3), is_simple, cal_float)
     Else
      divide_string = time_string(divide_string, _
         s(4), is_simple, cal_float)
      divide_string = divide_string(divide_string, _
         time_string(s(3), sq, False, cal_float), is_simple, cal_float)
     End If
   Else
   Call do_factor1(S2, f2(0), f2(1), f2(2), f2(3), 0) '因式分解除数
    sqr_no% = InStr(1, f2(0), "[", 0)
    If sqr_no% > 0 Then 'Asc(f2(0)) < 0 And f2(0) <> "'" And f2(0) <> "\" And Len(f2(0)) = 1 Then
     'tn% = -1
     'tn% = from_char_to_no(f2(0))
     sq = read_sqr_from_string(s(2), 0, s(6))
      For i% = 0 To 3
       If f1(i%) <> "1" Then
        For j% = 0 To 3
         If f2(j%) <> "1" Then
          If j% = 0 Then
           If f1(i%) = sq Then 'squre_root_string(tn%) Then
            f1(i%) = f2(0)
             f2(0) = "1"
           ElseIf f1(i%) = f2(0) Then
            f1(i%) = "1"
             f2(0) = "1"
           End If
         Else
          If f1(i%) = f2(j%) Then
           f1(i%) = "1"
            f2(j%) = "1"
          End If
         End If
       End If
     Next j%
     End If
    Next i%
    Else
      For i% = 0 To 3
       If f1(i%) <> "1" Then
       For j% = 0 To 3
         If f1(i%) = f2(j%) Then
         f1(i%) = "1"
         f2(j%) = "1"
       End If
     Next j%
     End If
    Next i%
    End If
     divide_string = time_string(f1(0), f1(1), False, cal_float)
     divide_string = time_string(divide_string, f1(2), False, cal_float)
     divide_string = time_string(divide_string, f1(3), is_simple, cal_float)
     If f(0) <> "1" And tn% >= 0 Then
     divide_string = time_string(divide_string, f2(0), False, cal_float)
     divide_string = divide_string(divide_string, number_string(tn%), is_simple, True)
     Else
      divide_string = divide_string(divide_string, f2(0), is_simple, False)
     End If
     divide_string = divide_string(divide_string, f2(1), False, cal_float)
     divide_string = divide_string(divide_string, f2(2), False, cal_float)
     divide_string = divide_string(divide_string, f2(3), is_simple, cal_float)
   End If
     Exit Function
ElseIf InStr(1, fs(1), "'", 0) > 0 Then
 If rational_string(S2, ts(0), ts(1)) Then
 '  ts(0) = combine_item_with_para(ts(0), S(4), True)
    divide_string = divide_string(time_string(s1, ts(0), False, cal_float), ts(1), _
       is_simple, cal_float)
  ' divide_string = simple_string_(divide_string, "", "")
    Exit Function
 Else
  divide_string = combine_divide_string(s1, S2)
   Exit Function
 End If
End If
End If
'****************************************************
If s(5) = "" Then '除数单项
 If s(2) = "" Then
   Call simple_multi_item(s(4), s(1), "", "", "", "", "", "", "", 2, "")
   Call simple_multi_para(s(3), s(0), "", "", "", "", "", "", "", 2, "", True, cal_float)
   s(3) = combine_item_with_para(s(3), s(4), True)
   s(0) = combine_item_with_para(s(0), s(1), True)
   If s(3) = "1" Then
    divide_string = s(0)
   Else
    If InStr(2, s(0), "-", 0) = 0 And InStr(2, s(0), "+", 0) = 0 And InStr(2, s(0), "#", 0) = 0 And _
        InStr(2, s(0), "@", 0) = 0 Then
     If InStr(2, s(3), "-", 0) = 0 And InStr(2, s(3), "+", 0) = 0 And InStr(2, s(3), "#", 0) = 0 And _
        InStr(2, s(3), "@", 0) = 0 Then
      divide_string = s(0) + "/" + s(3)
     Else
      divide_string = s(0) + "/" + "(" + s(3) + ")"
     End If
    Else
      If InStr(2, s(3), "-", 0) = 0 And InStr(2, s(3), "+", 0) = 0 And InStr(2, s(3), "#", 0) = 0 And _
        InStr(2, s(3), "@", 0) = 0 Then
       divide_string = "(" + s(0) + ")" + "/" + s(3)
      Else
       divide_string = "(" + s(0) + ")" + "/" + "(" + s(3) + ")"
      End If
    End If
   End If
 Else
 If read_string(s1, S_p_i) = False Then '分解除数
    GoTo divide_string_error
 End If
 If S_p_i.last_it = 1 Then
 If InStr(1, s(3), ".", 0) > 0 Or InStr(1, S_p_i.pA(0), ".", 0) > 0 Then
 S_p_i.pA(0) = divide_para(S_p_i.pA(0), s(3), is_simple, cal_float)
 s(3) = "1"
 'S_p_i.pa(0) = "1"
 Else
  If simple_multi_para(s(3), S_p_i.pA(0), "", "", "", "", "", "", "", 2, "", False, cal_float) = False Then
   divide_string = "F"
    Exit Function
  End If
 'End If
  If simple_multi_item(s(4), S_p_i.it(0), "", "", "", "", "", "", "", 2, "") = False Then
   divide_string = "F"
   Exit Function
  End If
   End If
  s(0) = time_string(S_p_i.pA(0), S_p_i.it(0), is_simple, cal_float)
 Else
 Call simple_multi_para_for_sp(S_p_i, fc(0))
 Call simple_multi_item_for_sp(S_p_i, fc(1))
 Call simple_multi_para(s(3), fc(0), "", "", "", "", "", "", "", 2, "", False, cal_float)
 Call simple_multi_item(s(4), fc(1), "", "", "", "", "", "", "", 2, "")
 ' Call simple_multi_para(s(3), S_p_i.pa(0), S_p_i.pa(1), S_p_i.pa(2), _
           S_p_i.pa(3), S_p_i.pa(4), S_p_i.pa(5), S_p_i.pa(6), _
             S_p_i.pa(7), S_p_i.last_it + 1, "")
  ' Call simple_multi_item(s(4), S_p_i.it(0), S_p_i.it(1), S_p_i.it(2), _
           S_p_i.it(3), S_p_i.it(4), S_p_i.it(5), S_p_i.it(6), _
             S_p_i.it(7), S_p_i.last_it + 1, "")
 fc(0) = time_string(fc(0), fc(1), is_simple, cal_float)
 s(0) = time_string(fc(0), _
   time_string(S_p_i.pA(0), S_p_i.it(0), False, False), is_simple, cal_float)
  End If
  For i% = 1 To S_p_i.last_it - 1
    s(5) = time_string(fc(0), _
      time_string(S_p_i.pA(i%), S_p_i.it(i%), False, cal_float), is_simple, cal_float)
    ts(4) = Mid$(s(5), 1, 1)
    If ts(4) = "-" Or s(4) = "+" Or ts(4) = "@" Or ts(4) = "#" Then
     s(0) = s(0) + s(5)
    Else
    s(0) = s(0) + "+" + s(5) 'time_string(fc(0), _
      combine_item_with_para(S_p_i.pa(i%), S_p_i.it(i%)), True, cal_float)
    End If
  Next i%
    s(3) = time_string(s(3), s(4), True, cal_float)
 If s(3) = "1" Then
  divide_string = s(0)
 Else
  If S_p_i.last_it = 1 Then
  divide_string = combine_divide_string(s(0), s(3))
  Else
  divide_string = "(" + s(0) + ")" + "/" + s(3)
  End If
 End If
 End If
 Else '分母是多项
   Call simple_string0(s1, ts(0))
    Call simple_string0(S2, ts(1))
 If InStr(1, s1, "F", 0) > 0 Or InStr(1, S2, "F", 0) > 0 Then
    divide_string = "F"
     Exit Function
 ElseIf ts(0) = ts(1) And ts(0) <> "1" Then
  ts(0) = "1"
   ts(1) = "1"
 End If
 If s1 = S2 Then
  s1 = "1"
   S2 = "1"
 End If
 Call simple_two_string_(s1, S2, "")
     If ts(0) <> "1" And ts(1) <> "1" Then
       Call simple_multi_string(ts(1), ts(0), "", "", "", "", "", "", "", 2, "", False, cal_float)
     End If
    s1 = time_string(s1, ts(0), False, cal_float)
     S2 = time_string(S2, ts(1), False, cal_float)
  For i% = 0 To 5
   s(i%) = ""
  Next i%
  is_t(0) = do_factor1(s1, pA(0), f(0), f(1), f(2), 0)
  is_t(1) = do_factor1(S2, pA(1), f(3), f(4), f(5), 0)
  Call simple_multi_string0(pA(1), pA(0), "", "", "", False)
If is_t(0) And is_t(1) Then
  For i% = 0 To 2
   For j% = 3 To 5
    If f(i%) = f(j%) Then
     f(i%) = "1"
     f(j%) = "1"
    End If
   Next j%
  Next i%
  s1 = time_string(f(0), f(1), False, cal_float)
   s1 = time_string(s1, f(2), False, cal_float)
   s1 = time_string(s1, pA(0), is_simple, cal_float)
   S2 = time_string(f(3), f(4), False, cal_float)
    S2 = time_string(S2, f(5), False, cal_float)
   S2 = time_string(S2, pA(1), is_simple, cal_float)
 If S2 = "1" Then
 divide_string = s1
 Exit Function
 Else
 ts(4) = add_brace(s1, "string")
 ts(5) = add_brace(S2, "string")
 If ts(4) = "1" Then
  If ts(5) = "1" Then
   divide_string = "1"
  Else
  divide_string = "1/" + ts(5)
  End If
 Else
  If ts(5) = "1" Then
   divide_string = ts(4)
  Else
  divide_string = ts(4) + "/" + ts(5)
  End If
 End If
 Exit Function
 End If
Else
 If InStr(1, S2, "'", 0) > 0 Then
 ts(4) = add_brace(s1, "")
 ts(5) = add_brace(S2, "root")
 If ts(4) = "1" Then
  If ts(5) = "1" Then
   divide_string = "1"
  Else
  divide_string = "1/" + ts(5)
  End If
 Else
  If ts(5) = "1" Then
   divide_string = ts(4)
  Else
  divide_string = ts(4) + "/" + ts(5)
  End If
 End If
 Else
  ts(4) = add_brace(s1, "")
  ts(5) = add_brace(S2, "all")
 If ts(4) = "1" Then
  If ts(5) = "1" Then
   divide_string = "1"
  Else
  divide_string = "1/" + ts(5)
  End If
 Else
  If ts(5) = "1" Then
   divide_string = ts(4)
  ElseIf ts(5) = "-1" Then
   divide_string = time_string("-1", ts(4), True, False)
  Else
  divide_string = ts(4) + "/" + ts(5)
  End If
 End If
 End If
 'Exit Function
End If
'Else
 ' Call add_brace_for_string(s1)
  '  Call add_brace_for_string(S2)
'divide_string = s1 + "/" + S2
 'Exit Function
'End If
End If
'End If
ElseIf t(0) = 3 And t(1) = 0 Then
If s(0) = S2 Then
divide_string = divide_string("1", s(1), is_simple, cal_float)
Else
If gcd_for_string(s(0), S2, "", fs(0), fs(1), True) Then
  s(0) = fs(0)
  S2 = fs(1)
End If
divide_string = divide_string(s(0), time_string(S2, s(1), False, cal_float), is_simple, cal_float)
End If
ElseIf t(0) = 0 And t(1) = 3 Then
If s1 = s(3) Then
Call remove_brace(s(4)) '约分
divide_string = s(4)
Else
If gcd_for_string(s1, s(3), "", fs(0), fs(1), True) Then
  s1 = fs(0)
  s(3) = fs(1)
End If
Call simple_two_string_(s(4), s(3), "")
'Call simple_two_string_(s1, s(3), "")
divide_string = divide_string(time_string(s(4), s1, False, cal_float), s(3), is_simple, cal_float)
End If
ElseIf t(0) = 3 And t(1) = 3 Then
If s(0) = s(3) Then
divide_string = divide_string(s(4), s(1), is_simple, cal_float)
ElseIf s(1) = s(4) Then
divide_string = divide_string(s(0), s(3), is_simple, cal_float)
Else
If gcd_for_string(s(0), s(3), "", fs(0), fs(1), True) Then
  s(0) = fs(0)
  s(3) = fs(1)
End If
If gcd_for_string(s(1), s(4), "", fs(0), fs(1), True) Then
  s(1) = fs(0)
  s(4) = fs(1)
End If
divide_string = divide_string(time_string(s(0), s(4), False, cal_float), _
       time_string(s(1), s(3), False, cal_float), False, cal_float)
End If
End If
If v_string_type = 1 Or v_string_type = 2 Then
  t(0) = string_type(divide_string, fs(0), s(0), s(1), s(2))
   If t(0) = 3 Then
    If InStr(1, s(0), "UU", 0) = 0 And InStr(1, s(0), "VV", 0) = 0 And InStr(1, s(0), "UV", 0) = 0 Then
     divide_string = "F"
      Exit Function
    End If
   ElseIf t(0) = 0 And v_string_type = 2 Then
     divide_string = "F"
      Exit Function
   End If
ElseIf v_string_type = 3 Then
  t(0) = string_type(divide_string, fs(0), s(0), s(1), s(2))
   If t(0) = 3 Then
    If InStr(1, s(1), "UU", 0) = 0 And InStr(1, s(1), "VV", 0) = 0 And InStr(1, s(1), "UV", 0) = 0 Then
     divide_string = "F"
      Exit Function
    End If
   End If
ElseIf v_string_type = 4 Then
  t(0) = string_type(divide_string, fs(0), s(0), s(1), s(2))
   If t(0) = 3 Then
    If InStr(1, s(1), "U", 0) > 0 Or InStr(1, s(1), "V", 0) > 0 Then
     divide_string = "F"
      Exit Function
    End If
   End If
End If
If is_simple = True Then
divide_string = simple_string_(divide_string, "", "", cal_float)
End If
Exit Function
divide_string_error:
divide_string = "F"
End Function
Public Function do_factor(ByVal pA1$, ByVal pA2$, ByVal pa3$, _
   ByVal pa4$, ByVal it1$, ByVal it2$, ByVal it3$, ByVal it4$, _
      ByVal n%, main_para$, f1$, f2$, f3$) As Boolean
'f1$ it$,f2$,f3$ 既约因子
Dim ch1(2) As String
'Dim ch2(2) As String * 1
'Dim ch3(3) As String * 1
'Dim ch4(3) As String * 1
Dim ch(1) As String
Dim ts(3) As String
Dim t_item(1) As String
Dim t_pa(1) As String
Dim order As Byte
main_para$ = "1"
f1$ = "1"
f2$ = "1"
f3$ = "1"
On Error GoTo do_factor_error
do_factor = True
Call simple_multi_item(it1$, it2$, it3$, it4$, "", "", "", "", "", n%, f1$)
'f1%是后比前
Call simple_multi_para0(pA1$, pA2$, pa3$, pa4$, main_para$)
 If n% = 1 Then
  main_para$ = pA1$
   f1$ = it1$
    do_factor = True
    Exit Function
 ElseIf n% = 2 Then
   If root_item(it1$, 2, t_item(0)) And root_item(it2$, 2, t_item(1)) Then
    ch1(0) = t_item(0)
    ch1(1) = t_item(1)
    order = 2
     GoTo do_factor_mark1
   End If
 ElseIf n% = 3 Then
   If root_item(it1$, 2, t_item(0)) And root_item(it3$, 2, t_item(1)) Then
    If time_string(t_item(0), t_item(1), True, False) = it2$ Then
     ch1(0) = t_item(0)
     'ch1(1) = it2$
     ch1(1) = t_item(1)
     order = 2
     GoTo do_factor_mark1
    End If
   End If
 ElseIf n% = 4 Then
 End If
'order = max(order, Len(it1$))
'order = max(order, Len(it2$))
'order = max(order, Len(it3$))
'order = max(order, Len(it4$))
'If order > 3 Or order = 1 Then
 f2$ = combine_string_from_para_item(pA1$, pA2$, pa3$, pa4$, _
          it1$, it2$, it3$, it4$, n%)
  do_factor = True
   Exit Function
'End If
 'ch1(0) = Mid$(it1$, 1, 1)
'If Len(it1$) = 2 Then
' ch1(1) = Mid$(it1$, 2, 1)
'ElseIf Len(it1$) = 3 Then
' ch1(1) = Mid$(it1$, 2, 1)
' ch1(2) = Mid$(it1$, 3, 1)
'ElseIf Len(it1$) > 3 Then
'  f2$ = combine_string_from_para_item(pa1$, pa2$, pa3$, pa4$, _
'          it1$, it2$, it3$, it4$, n%)
'  do_factor = True
'   Exit Function
'End If
'******************
' ch2(0) = Mid$(it2$, 1, 1)
'If Len(it1$) = 2 Then
' ch2(1) = Mid$(it2$, 2, 1)
'ElseIf Len(it2$) = 3 Then
' ch2(1) = Mid$(it2$, 2, 1)
' ch2(2) = Mid$(it2$, 3, 1)
'ElseIf Len(it2$) > 3 Then
'  f2$ = combine_string_from_para_item(pa1$, pa2$, pa3$, pa4$, _
          it1$, it2$, it3$, it4$, n%)
'  do_factor = True
'   Exit Function
'End If
'*****************
' ch3(0) = Mid$(it3$, 1, 1)
'If Len(it1$) = 2 Then
' ch3(1) = Mid$(it3$, 2, 1)
'ElseIf Len(it1$) = 3 Then
' ch3(1) = Mid$(it3$, 2, 1)
' ch3(2) = Mid$(it3$, 3, 1)
'ElseIf Len(it3$) > 3 Then
' f2$ = combine_string_from_para_item(pa1$, pa2$, pa3$, pa4$, _
          it1$, it2$, it3$, it4$, n%)
'  do_factor = True
'   Exit Function
'End If
'***************
' ch4(0) = Mid$(it4$, 1, 1)
'If Len(it1$) = 2 Then
' ch4(1) = Mid$(it4$, 2, 1)
'ElseIf Len(it1$) = 3 Then
' ch4(1) = Mid$(it4$, 2, 1)
' ch4(2) = Mid$(it4$, 3, 1)
'ElseIf Len(it4$) > 3 Then
' f2$ = combine_string_from_para_item(pa1$, pa2$, pa3$, pa4$, _
          it1$, it2$, it3$, it4$, n%)
'  do_factor = True
'   Exit Function
'End If
'main_para$ = pa1$
'pa1$ = 1
'pa2$ = divide_string(pa2$, main_para$)
'pa3$ = divide_string(pa3$, main_para$)
'pa4$ = divide_string(pa4$, main_para$)
'*********************************************
'If order = 1 Then '一次
'If f1$ = "1" Then
 'f1$ = combine_string_from_para_item(pa1$, pa2$, "", "", _
          ch1(0), ch2(0), "", "", 2)
'Else
' f2$ = combine_string_from_para_item(pa1$, pa2$, "", "", _
          ch1(0), ch2(0), "", "", 2)
'End If
   ' do_factor = True
    '  Exit Function
do_factor_mark1:
If order = 2 Then
If n% = 2 Then  '平方差
      If (Mid$(pA2$, 1, 1) = "-" And Mid$(pA1$, 1, 1) <> "-") Or _
        (Mid$(pA2$, 1, 1) <> "-" And Mid$(pA1$, 1, 1) = "-") Then
      If Mid$(pA1$, 1, 1) = "-" Then
       ts(0) = Mid$(pA1$, 2, Len(pA1$) - 1)
      Else
       ts(0) = pA1$
      End If
      If Mid$(pA2$, 1, 1) = "-" Then
       ts(1) = Mid$(pA2$, 2, Len(pA2$) - 1)
      Else
       ts(1) = pA2$
      End If
'88888888888888888888888888888888888888
      ts(0) = sqr_string(ts(0), False, False)
      If InStr(1, ts(0), "F", 0) > 0 Then
       do_factor = False
         Exit Function
      End If
      ts(1) = sqr_string(ts(1), False, False)
      If InStr(1, ts(1), "F", 0) > 0 Then
      do_factor = False
      Exit Function
      End If
     f2$ = combine_string_from_para_item(ts(0), ts(1), "", "", _
           ch1(0), ch1(1), "", "", 2)
      If Mid$(pA2$, 1, 1) = "-" Then
       f3$ = f2$
       f2$ = combine_string_from_para_item(ts(0), time_string("-1", ts(1), False, False) _
           , "", "", ch1(0), ch1(1), "", "", 2)
      Else
       f3$ = combine_string_from_para_item(time_string(ts(0), "-1", False, False), ts(1) _
           , "", "", ch1(0), ch1(1), "", "", 2)
      End If
     Else ' 平方和
     f2$ = combine_string_from_para_item(pA1$, pA2$, pa3$, pa4$, _
          it1$, it2$, it3$, it4$, n%)
     '  If f1$ = "1" Then
      '  f1$ = f2$
       '  f2$ = "1"
      ' End If
        do_factor = True
       Exit Function
    End If
 ElseIf n% = 3 Then '二次三项
'  If ch3(0) = ch3(1) And _
       ((ch2(0) = ch1(0) And ch2(1) = ch3(0)) Or _
              ch1(0) = "1" And ch2(0) = ch3(0)) Then
 '    ch(0) = ch1(0)
  '    ch(1) = ch3(0)
 If pA1$ = "1" And pA2$ = "2" And pa3$ = "1" Then
     f2$ = ch1(0) + "+" + ch1(1)
      f3$ = ch1(0) + "+" + ch1(1)
       'If f1$ = "1" Then
        'f1$ = f3$
         'f3$ = ""
       'End If
    do_factor = True
   Exit Function
    ElseIf pA1$ = "1" And pA2$ = "-2" And pa3$ = "1" Then
     f2$ = ch1(0) + "-" + ch1(1)
      f3$ = ch1(0) + "-" + ch1(1)
     do_factor = True
       Exit Function
    Else '不是平方和的公式
     ts(0) = time_string(pA2$, pA2$, False, False)
      ts(0) = minus_string(ts(0), time_string("4", time_string(pA1$, pa3$, False, False), _
                False, False), False, False)
     If is_less_than(ts(0), "0") Then
       f2$ = combine_string_from_para_item(pA1$, pA2$, pa3$, pa4$, _
          it1$, it2$, it3$, it4$, n%)
         '  If f1$ = "1" Then
          '  f1$ = f2$
           '  f2$ = "1"
          ' End If
           do_factor = True
       Exit Function
    Else
      ts(0) = sqr_string(ts(0), False, False)
       If InStr(1, ts(0), "F", 0) > 0 Or InStr(1, ts(0), "'", 0) > 0 Then
        do_factor = False
         Exit Function
       End If
       
       ts(1) = add_string(pA2$, ts(0), False, False)
        ts(2) = minus_string(pA2$, ts(0), False, False)
         ts(1) = divide_string(ts(1), time_string("2", pA1$, False, False), False, False)
         ts(2) = divide_string(ts(2), time_string("2", pA1$, False, False), False, False)
      If pA1$ = "1" Then
      f2$ = combine_string_from_para_item("1", ts(1), "", "", _
            ch1(0), ch1(1), "", "", 2)
      f3$ = combine_string_from_para_item("1", ts(2), "", "", _
           ch1(0), ch1(1), "", "", 2)
      Else
      ts(0) = "1"
      ts(3) = "1"
      Call simple_multi_para(ts(0), ts(1), "", "", "", "", "", _
         "", "", 2, "", False, False)
      Call simple_multi_para(ts(3), ts(2), "", "", "", "", "", _
         "", "", 2, "", False, False)
      f2$ = combine_string_from_para_item(ts(0), ts(1), "", "", _
            ch1(0), ch1(1), "", "", 2)
      f3$ = combine_string_from_para_item(ts(3), ts(2), "", "", _
           ch1(0), ch1(1), "", "", 2)
         
      End If
       'If f1$ = "1" Then
        'f1$ = f3$
         'f3$ = "1"
      ' End If
       do_factor = True
        Exit Function
     End If ' 判别式大于零
 End If
Else
 f2$ = combine_string_from_para_item(pA1$, pA2$, pa3$, pa4$, _
          it1$, it2$, it3$, it4$, n%)
  'If f1$ = "1" Then
   'f1$ = f2$
    'f2$ = "1"
  'End If
  do_factor = True
   Exit Function
   End If
  Else
 f2$ = combine_string_from_para_item(pA1$, pA2$, pa3$, pa4$, _
          it1$, it2$, it3$, it4$, n%)
  'If f1$ = "1" Then
   'f1$ = f2$
    'f2$ = "1"
  'End If
  do_factor = True
   Exit Function
  End If
 'End If
'ElseIf order = 3 Then
'If n% = 2 Then
Exit Function
do_factor_error:
do_factor = False
'ElseIf n% = 4 Then

End Function

Public Function dmin_0(X%, Y%) As Integer
If X% <= Y% Or (Y% = 0 And X% > 0) Then
dmin_0 = X%
ElseIf X% > Y% Or (X% = 0 And Y% > 0) Then
dmin_0 = Y%
Else
dmin_0 = 0
End If


End Function


Private Function value_para(ByVal p As String) As String
Dim t%, tn%, i%
Dim ch As String
Dim tp(3) As Integer
Dim s(3) As String
Dim v As Variant
'Dim fs As String
On Error GoTo value_para_error
If p = "0" Then
 value_para = "0"
  Exit Function
End If
Call remove_brace(p)
t% = para_type(p, "", s(0), s(1), s(2))
  If Mid$(s(0), 1, 1) = "@" Then
   s(0) = "-" + Mid$(s(0), 2, Len(s(0)) - 1)
  ElseIf Mid$(s(0), 1, 1) = "#" Then
   s(0) = Mid$(s(0), 2, Len(s(0)) - 1)
  End If
If t% = 2 Then
If Mid$(p, 1, 1) = "@" Then
value_para = "-" + Mid$(p, 2, Len(p) - 1)
Else
value_para = p
End If
ElseIf t% = 0 Then
 '整数和消数
 For i% = 1 To Len(p)
   ch = Mid$(p, i%, 1)
    If ch >= "A" And ch <> "'" Then
       value_para = "F"
        Exit Function
    End If
 Next i%
If s(1) = "" Or s(1) = "1" Then
  value_para = p
Else
  ch = Mid$(s(0), 1, 1)
  v = val_(value_para(s(1)))
  If v >= 0 Then
   value_para = str_(val_(value_para(s(0))) * sqr(v) + _
                        val_(value_para(s(2))))
  Else
   value_para = "F"
    Exit Function
  End If
 End If
ElseIf t% = 3 Then
 Call remove_brace(s(0))
   Call remove_brace(s(1))
    s(0) = value_para(s(0))
    s(1) = value_para(s(1))
    If s(0) <> "F" And s(1) <> "F" Then
     value_para = str_(val_(s(0) / _
       val_(s(1))))
    Else
     value_para = "F"
    End If
 '分数
ElseIf t% = 1 Then
  '根式
'If para_type(s(1), "", "", "") <> 1 Then
s(1) = value_para(s(1))
v = val_(value_para(s(1)))
If v <> "F" Then
value_para = str_(val_(value_para(s(0))) * sqr(val_(s(1))))
Else
value_para_error:
value_para = "F"
Exit Function
End If
End If
End Function


Private Function string_from_para_item(p_i As para_item_type) As String
Dim i%
If p_i.last_it = 0 Then
string_from_para_item = "0"
Else
For i% = 0 To p_i.last_it - 1
string_from_para_item = combine_para_or_string_for_add( _
 string_from_para_item, time_string(p_i.pA(i%), _
   p_i.it(i%), True, False), "string")
Next i%
End If
End Function
Public Function string_type_(ByVal s As String, s0 As String, s1 As String, _
                       S2 As String, s3 As String) As Integer
Dim i%, start%, f_brace%, e_brace%
Dim ch$
 If InStr(1, s, "/", 0) > 0 Then
  string_type_ = string_type(s, s0, s1, S2, s3)
 Else
 start% = InStr(1, s, "(", 0)
   If start% = 1 Then
      f_brace% = InStr(2, s, ")", 0)
       e_brace% = InStr(f_brace + 1, s, "(", 0)
        If e_brace% = 0 Then
            e_brace% = Len(s) + 1
        End If
         f_brace% = f_brace% + 1
   Else
         f_brace% = 1
         e_brace% = start% - 1
   End If
   For i% = f_brace% To e_brace%
      ch$ = Mid$(s, i%, 1)
       If ch$ = "+" Or ch$ = "-" Or ch$ = "#" Or ch$ = "@" Then
        GoTo string_type_mark10
       End If
   Next i%
   start% = start% + 1
   If is_right_an_item(s, start%) Then
       string_type_ = 2
        s1$ = Mid$(s, 1, start% - 2)
        If Mid$(s, start% - 1, 1) = "*" Then
           S2$ = Mid$(s, start%, Len(s) - start% + 1)
        Else
           S2$ = Mid$(s, start% - 1, Len(s) - start% + 2)
        End If
         string_type_ = 2
          Exit Function
   End If
string_type_mark10:
 string_type_ = string_type(s, s0, s1, _
       S2, s3)
End If
End Function
Public Function string_type(ByVal s As String, s0 As String, s1 As String, _
                            S2 As String, s3 As String) As Integer
 '0整式
 '1 ,
 '2
 '3 ()/()
Dim ts0(3) As String
Dim ts As String
Dim ts1 As String
Dim f_brace%
Dim e_brace%
Dim start%
 Dim p%, n%, m%, i%, j%, k%, b1%, b2%, k1%, k2%
  Dim tn(1) As Integer
Call remove_brace(s)
If s = "" Or s = "0" Then
 s0 = "0"
  s1 = "0"
   S2 = "0"
    s3 = ""
     string_type = 0
      Exit Function
 End If
p% = 0
ts = s
Call unsimple_string(ts)
Call get_brace_pair(ts, 1, f_brace%, e_brace%)
If f_brace% > 0 Then
 start% = e_brace%
Else
   f_brace% = 1000
   e_brace% = 1000
End If
n% = InStr(1, ts, "/", 0) '分式
If n% = 0 Then
   Call read_item_from_string(ts, s0, s3)
   Call read_para_from_item(s0, s1, S2)
ts0(1) = s3
Do
  Call read_item_from_string(ts0(1), ts0(0), ts0(1)) '读出后一项
     Call read_para_from_item(ts0(0), ts0(2), ts) '分离系数
      If S2 = ts And ts0(2) <> "" Then '是同类项
        s3 = ts0(1)
         If Mid$(ts0(0), 1, 1) = "+" Or Mid$(ts0(0), 1, 1) = "1" Or Mid$(ts0(0), 1, 1) = "#" Or _
              Mid$(ts0(0), 1, 1) = "@" Then
             s0 = s0 & ts0(0)
         Else
             s0 = s0 & "+" & ts0(0)
         End If
         s1 = add_para(s1, ts0(2), False, False)
      End If
Loop Until ts <> S2 Or ts0(1) = ""
      string_type = 0
ElseIf n% > 0 Then
 s1 = Mid$(s$, 1, n% - 1)
  S2 = Mid$(s$, n% + 1, Len(s$) - n%)
   string_type = 3
End If
 If s1 = "" Or s1 = "+" Or s1 = "#" Then
  s1 = "1"
 ElseIf s1 = "-" Or s1 = "@" Then
  s1 = "-1"
 End If
 'If S2 = "" Then
 ' S2 = "1"
 'End If
End Function
Private Sub sq_root(n1&, n2&, n3&)
Dim q&
Dim tn!
q& = CLng(sqr(n1&))
For n2& = q& To 1 Step -1
tn! = n1& / (n2& * n2&)
 n3& = n1& / (n2& * n2&)
 If tn! = n3& Then
 Exit Sub
 End If
Next n2&
End Sub

Public Function is_int_divide(n1%, n2%, n3%) As Boolean
n3% = n1% / n2%
If n1% = n2 * n3% Then
 is_int_divide = True
End If

End Function

Private Function simple_multi_int(n1&, n2&, n3&, n4&, n5&, n6&, n7&, n8&, n9&, n%, n0&) As Boolean
Dim i%, j%, k%
Dim tn(8) As Long
Dim tn1(8) As Long
Dim t_n&
On Error GoTo simple_multi_int_error
simple_multi_int = True
If n1 <> 0 Then
n0& = n1&
Else
n0& = 1
End If
'k% = n1%
If n1& > 0 Then
tn(0) = n1&
tn(1) = n2&
tn(2) = n3&
tn(3) = n4&
tn(4) = n5&
tn(5) = n6&
tn(6) = n7&
tn(7) = n8&
tn(8) = n9&
ElseIf n1& < 0 Then
tn(0) = -n1&
tn(1) = -n2&
tn(2) = -n3&
tn(3) = -n4&
tn(4) = -n5&
tn(5) = -n6&
tn(6) = -n7&
tn(7) = -n8&
tn(8) = -n9&
Else
 If n1& <> 0 Then
 n0& = n0& / n1&
 End If
 Exit Function
End If
t_n& = tn(0)
 For j% = 1 To n%
  t_n& = l_gcd(t_n&, tn(j%))
 Next j%
n1& = tn(0) / t_n&
n2& = tn(1) / t_n&
n3& = tn(2) / t_n&
n4& = tn(3) / t_n&
n5& = tn(4) / t_n&
n6& = tn(5) / t_n&
n7& = tn(6) / t_n&
n8& = tn(7) / t_n&
n9& = tn(8) / t_n&
If n1& <> 0 Then
n0& = n0& / n1&
End If
Exit Function
simple_multi_int_error:
simple_multi_int = False
End Function

Private Function remove_space(s As String) As String
Dim p%
Dim ts As String
ts = s
Do
 p% = InStr(1, ts, " ", 0)
If p% = 0 Then
 remove_space = ts
  Exit Function
Else
  ts = Mid$(ts, 1, p% - 1) + Mid$(ts, p% + 1, Len(ts) - p%)
End If
Loop
End Function

Public Function link_para(p1 As String, p2 As String) As String
If p1 = "" Or p1 = "0" Then
 link_para = p2
ElseIf p2 = "" Or p2 = "0" Then
 link_para = p1
Else
 link_para = p1 + "#" + p2
End If
End Function
Public Function combine_para_for_item(pA As String, its As String) As String
'形成项
If InStr(1, pA, "F", 0) > 0 Or InStr(1, its, "F", 0) > 0 Then
  combine_para_for_item = "F"
   Exit Function
ElseIf pA = "" Or pA = "0" Or its = "0" Then
 combine_para_for_item = "0"
Else
 If its = "1" Or its = "" Then
  'If Mid$(pa, 1, 1) = "@" Then
   ''combine_para_for_item = "-" + Mid$(pa, 2, Len(pa) - 1)
  'Else
   combine_para_for_item = pA
  'End If
 Else
 ' its = add_brace_for_string_or_para(its, "string")
   pA = add_brace(pA, "")
  If pA = "1" Then
   combine_para_for_item = "'" + its
  ElseIf pA = "-1" Or pA = "@1" Then
   combine_para_for_item = "@" + "'" + its
  Else
   combine_para_for_item = pA + "'" + its
  End If
 End If
End If
End Function


Private Function combine_para_or_string_for_add(ByVal s1$, ByVal S2$, ty As String) As String
Dim sg$
If InStr(1, s1$, "F", 0) > 0 Or InStr(1, S2$, "F", 0) > 0 Then
 combine_para_or_string_for_add = "F"
  Exit Function
End If
sg$ = Mid$(S2$, 1, 1)
 If s1$ = "0" Or S2$ = "0" Then
 combine_para_or_string_for_add = "0"
 ElseIf (s1$ = "0" Or s1$ = "") And S2$ <> "" And S2$ <> "0" Then
 combine_para_or_string_for_add = S2$
 ElseIf (S2$ = "0" Or S2$ = "") And s1$ <> "" And s1$ <> "0" Then
 combine_para_or_string_for_add = s1$
 ElseIf ty = "string" Then
 If sg = "-" Then
 combine_para_or_string_for_add = s1$ + S2$
 Else
 combine_para_or_string_for_add = s1$ + "+" + S2$
 End If
 ElseIf ty = "para" Then
 If sg = "@" Or sg = "#" Then
 combine_para_or_string_for_add = s1$ + S2$
 ElseIf sg = "-" Then
 combine_para_or_string_for_add = s1$ + "@" + Mid$(S2$, 2, Len(S2) - 1)
 ElseIf sg = "+" Then
 combine_para_or_string_for_add = s1$ + "#" + Mid$(S2$, 2, Len(S2) - 1)
 Else
 combine_para_or_string_for_add = s1$ + "#" + S2$
 End If
 End If
End Function


Private Function simple_multi_item(i1$, i2$, i3$, I4$, I5$, I6$, I7$, I8$, I9$, n%, I0$) As Boolean
Dim i%, j%, st%, l_of_ch%
Dim ti(8) As String
Dim ts1(8) As String
Dim ts2(8) As String
Dim ch As String
Dim ty As Byte '=0=1 根号
Dim tn(8) As Integer
Dim t_I0$
On Error GoTo simple_multi_item_error
I0$ = "1"
If n% < 2 Then
 Exit Function
End If
t_I0$ = i1$
ti(0) = i1$
ti(1) = i2$
ti(2) = i3$
ti(3) = I4$
ti(4) = I5$
ti(5) = I6$
ti(6) = I7$
ti(7) = I8$
ti(8) = I9$
For i% = 0 To n% - 1
 If ti(i%) = "1" Then
 I0$ = "1"
 simple_multi_item = True
  Exit Function
 End If
Next i%
i% = 1
Do While Len(ti(0)) >= i% And ti(0) <> "1"
 st% = i%
  ch = Mid$(ti(0), i%, 1) '取出一个字
   If ch = "[" Then
    Call read_sqr_no_from_string(ti(0), i%, i%, ch)
   ElseIf ch = "'" Then
    Call read_sqr_string_from_string(ti(0), i%, i%, "", ch, "")
    ty = 1
   End If
'****************************************************
For j% = 1 To n% - 1
 If ty = 1 Then
  Call read_sqr_string_from_string(ti(j%), 1, 0, "", ts1(j%), ts2(j%))
   If ts1(j%) <> ch Then
    i% = i% + 1
     GoTo simple_multi_item_mark1
   End If
 Else
 tn(j%) = InStr(1, ti(j%), ch, 0)
  If tn(j%) = 0 Then
   i% = i% + 1
    GoTo simple_multi_item_mark1
  Else
   If tn(j%) > 1 Then
     If Mid$(ti(j%), tn(j%) - 1, 1) = "'" Then
      tn(j%) = 0
       i% = i% + 1
        GoTo simple_multi_item_mark1
     End If
   End If
  End If
 End If
Next j%
'*****************************************************************************
'有公因子ch
l_of_ch = Len(ch)
ti(0) = Mid$(ti(0), 1, st% - 1) + Mid$(ti(0), st% + l_of_ch, Len(ti(0)) - st% - l_of_ch + 1)
If ti(0) = "" Then
 ti(0) = "1"
End If
For j% = 1 To n% - 1
 If ty = 0 Then
 ti(j%) = Mid$(ti(j%), 1, tn(j%) - 1) + Mid$(ti(j%), tn(j%) + l_of_ch, Len(ti(j%)) - tn(j%) - l_of_ch + 1)
 If ti(j%) = "" Then
  ti(j%) = "1"
 End If
 Else
  ti(j%) = ts2(j%)
 End If
Next j%
i% = st%
simple_multi_item_mark1:
Loop
simple_multi_item_mark2:
i1$ = ti(0)
i2$ = ti(1)
i3$ = ti(2)
I4$ = ti(3)
I5$ = ti(4)
I6$ = ti(5)
I7$ = ti(6)
I8$ = ti(7)
I9$ = ti(8)
I0$ = divide_item(t_I0$, i1$, "", "")
simple_multi_item = True
 Exit Function
simple_multi_item_error:
 simple_multi_item = False
End Function

Public Function simple_multi_string(s1$, S2$, s3$, S4$, S5$, S6$, S7$, S8$, S9$, n%, s0$, _
               is_simple As Boolean, cal_float As Boolean) As Boolean '同时简化几个表达式
Dim ts(8) As String
Dim para(8) As String
Dim item(8) As String
Dim ls As String
Dim t As Integer
Dim i%
On Error GoTo simple_multi_string_error
simple_multi_string = True
s0$ = "1"
If n% < 2 Then
 Exit Function
ElseIf n = 2 Then
 Call string_type(s1$, "", para(1), item(1), ts(1))
  Call string_type(S2$, "", para(2), item(2), ts(2))
   Call simple_multi_para(para(1), para(2), "", "", "", "", "", "", "", 2, para(0), True, False)
    Call simple_multi_item(item(1), item(2), "", "", "", "", "", "", "", 2, item(0))
     s1$ = combine_item_with_para(para(1), item(1), True)
      S2$ = combine_item_with_para(para(2), item(2), True)
       s0$ = combine_item_with_para(para(0), item(0), True)
Else
ts(0) = s1$
ts(1) = S2$
ts(2) = s3$
ts(3) = S4$
ts(4) = S5$
ts(5) = S6$
ts(6) = S7$
ts(7) = S8$
ts(8) = S9$
s0$ = s1$
For i% = 0 To n% - 1
t = string_type(ts(i%), "", para(i%), item(i%), ls)
 If ls <> "" And t = 0 Then
  Exit Function
 End If
Next i%
Call simple_multi_para(para(0), para(1), para(2), para(3), _
 para(4), para(5), para(6), para(7), _
  para(8), n%, "", True, cal_float)
Call simple_multi_item(item(0), item(1), item(2), item(3), _
 item(4), item(5), item(6), item(7), _
  item(8), n%, "")
For i% = 0 To n% - 1
 If item(i%) = "1" Or item(i%) = "" Then
  ts(i%) = para(i%)
 ElseIf para(i%) = "1" Or para(i%) = "" Then
  ts(i%) = item(i%)
 ElseIf para(i%) = "0" Then
  ts(i%) = "0"
 Else
  ts(i%) = para(i%) + "*" + item(i%)
 End If
Next i%
  s0$ = divide_string(s0$, ts(0), is_simple, False)
  If InStr(1, s0$, "U", 0) = 0 And InStr(1, s0$, "V", 0) = 0 Then
  s1$ = ts(0)
  S2$ = ts(1)
  s3$ = ts(2)
  S4$ = ts(3)
  S5$ = ts(4)
  S6$ = ts(5)
  S7$ = ts(6)
  S8$ = ts(7)
  S9$ = ts(8)
  Else
  s0$ = "1"
  End If
End If
Exit Function
simple_multi_string_error:
simple_multi_string = False
End Function

Private Function simple_multi_para(p1$, p2$, p3$, p4$, p5$, p6$, p7$, p8$, _
             p9$, n%, p0$, is_simple As Boolean, cal_float As Boolean) As Boolean
Dim i%, j%
Dim s1$
Dim S2$
Dim s3$
Dim p_i(8) As para_item_type
Dim lgc As Long
Dim g_c As String
Dim tp(8) As String
Dim tp1(8) As String
Dim tp2(8) As String
Dim fac(8) As String
Dim tn(9) As Long
Dim tr(9) As Long
p0$ = p1$
Dim ty As Byte
On Error GoTo simple_multi_para_error
If n% < 2 Then
simple_multi_para = True
 Exit Function
ElseIf cal_float = True Then
p1$ = "1"
p2$ = divide_para(p2$, p0$, is_simple, cal_float)
p3$ = divide_para(p3$, p0$, is_simple, cal_float)
p4$ = divide_para(p4$, p0$, is_simple, cal_float)
p5$ = divide_para(p5$, p0$, is_simple, cal_float)
p6$ = divide_para(p6$, p0$, is_simple, cal_float)
p7$ = divide_para(p7$, p0$, is_simple, cal_float)
p8$ = divide_para(p8$, p0$, is_simple, cal_float)
p9$ = divide_para(p9$, p0$, is_simple, cal_float)
   simple_multi_para = True
     Exit Function
End If
tp(0) = p1$
tp(1) = p2$
tp(2) = p3$
tp(3) = p4$
tp(4) = p5$
tp(5) = p6$
tp(6) = p7$
tp(7) = p8$
tp(8) = p9$
For i% = 0 To n% - 1
If InStr(1, p1$, ".", 0) > 0 Then
 For j% = 1 To n% - 1
  tp(j%) = divide_para(tp(j%), tp(0), True, True)
 Next j%
 p0$ = p1$
 p1$ = "1"
 p2$ = tp(1)
 p3$ = tp(2)
 p4$ = tp(3)
 p5$ = tp(4)
 p6$ = tp(5)
 p7$ = tp(6)
 p8$ = tp(7)
 p9$ = tp(8)
 simple_multi_para = True
 Exit Function
End If
Next i%
For i% = 0 To n% - 1
 If para_type(tp(i%), "", s1$, S2$, s3$) = 3 Then
  For j% = 0 To n% - 1
   If j% <> i% Then
    tp(j%) = time_para(tp(j%), S2$, is_simple, cal_float)
   Else
    tp(j%) = s1$
   End If
  Next j%
 End If
Next i%
If Mid$(tp(0), 1, 1) = "@" Or Mid$(tp(0), 1, 1) = "-" Then
For i% = 0 To n% - 1
tp(i%) = time_para("@1", tp(i%), is_simple, cal_float)
Next i%
End If
For i% = 0 To n% - 1
Call read_para(tp(i%), p_i(i%))
If i% = 0 Then
tn(0) = val_(p_i(0).pA(0))
tn(1) = val_(p_i(0).it(0))
End If
For j% = 0 To p_i(i%).last_it - 1
tn(0) = l_gcd(tn(0), val_(p_i(i%).pA(j%)))
tn(1) = l_gcd(tn(1), val_(p_i(i%).it(j%)))
Next j%
Next i%
For i% = 0 To n% - 1
 For j% = 0 To p_i(i%).last_it - 1
  p_i(i%).pA(j%) = str_(val_(p_i(i%).pA(j%)) / tn(0))
   p_i(i%).it(j%) = str_(val_(p_i(i%).it(j%)) / tn(1))
Next j%
  tp(i%) = para_from_para_item(p_i(i%))
Next i%
 p0 = divide_para(p0, tp(0), False, False)
p1$ = tp(0)
p2$ = tp(1)
p3$ = tp(2)
p4$ = tp(3)
p5$ = tp(4)
p6$ = tp(5)
p7$ = tp(6)
p8$ = tp(7)
p9$ = tp(8)
simple_multi_para = True
Exit Function
simple_multi_para_error:
simple_multi_para = False
End Function

Private Function simple_multi_para0(p1$, p2$, p3$, p4$, _
                    fa As String) As Boolean
Dim i% ' tr_p1% 扩大， tr_p2%
Dim tn(8) As Long
Dim tp(3) As String
Dim it(3) As String
Dim t_p As String
Dim s1$, S2$, ls$
Dim fr&
Dim gcf As Integer
Dim ty(3) As Byte
Dim tfa As String
On Error GoTo simple_multi_para0_error
If p1$ = "" Then
   p1$ = "0"
End If
If p2$ = "" Then
   p2$ = "0"
End If
If p3$ = "" Then
   p3$ = "0"
End If
If p4$ = "" Then
   p4$ = "0"
End If
simple_multi_para0 = True
If fa = "" Then
 fa = "1"
End If
If para_type(p1$, "", tp(0), tp(1), "") = 3 Then
   p1$ = tp(0)
   p2$ = time_para(p2$, tp(1), False, False)
   p3$ = time_para(p3$, tp(1), False, False)
   p4$ = time_para(p4$, tp(1), False, False)
   fa = divide_string(fa, tp(1), False, False)
   Call simple_multi_para0(p1$, p2$, p3$, p4$, fa)
   Exit Function
End If
If para_type(p2$, "", tp(0), tp(1), "") = 3 Then
   p2$ = tp(0)
   p1$ = time_para(p1$, tp(1), False, False)
   p3$ = time_para(p3$, tp(1), False, False)
   p4$ = time_para(p4$, tp(1), False, False)
   fa = divide_string(fa, tp(1), False, False)
   Call simple_multi_para0(p1$, p2$, p3$, p4$, fa)
   Exit Function
End If
If para_type(p3$, "", tp(0), tp(1), "") = 3 Then
   p3$ = tp(0)
   p2$ = time_para(p2$, tp(1), False, False)
   p1$ = time_para(p1$, tp(1), False, False)
   p4$ = time_para(p4$, tp(1), False, False)
   fa = divide_string(fa, tp(1), False, False)
   Call simple_multi_para0(p1$, p2$, p3$, p4$, fa)
   Exit Function
End If
If para_type(p4$, "", tp(0), tp(1), "") = 3 Then
   p4$ = tp(0)
   p2$ = time_para(p2$, tp(1), False, False)
   p3$ = time_para(p3$, tp(1), False, False)
   p1$ = time_para(p1$, tp(1), False, False)
   fa = divide_string(fa, tp(1), False, False)
   Call simple_multi_para0(p1$, p2$, p3$, p4$, fa)
   Exit Function
End If
tp(0) = p1$
tp(1) = p2$
tp(2) = p3$
tp(3) = p4$
   t_p = tp(0)
'*********
For i% = 0 To 3
If para_type(tp(i%), "", s1$, S2$, ls$) = 0 And ls$ = "" Then
     tn(2 * i% + 1) = val_(s1$)
     If tn(2 * i% + 1) <> 0 Then
     tn(2 * i%) = 1
     Else
     tn(2 * i%) = 0
     End If
ElseIf para_type(tp(i%), "", s1$, S2$, ls$) = 1 Then
     tn(2 * i% + 1) = val_(s1$)
      tn(2 * i%) = val_(S2)
    If tn(2 * i%) = 0 Then
     Exit Function
    End If
Else
 Exit Function
End If
Next i%
  Call simple_multi_int(tn(0), tn(2), tn(4), tn(6), 0, 0, 0, 0, 0, 4, fr)
  Call simple_multi_int(tn(1), tn(3), tn(5), tn(7), 0, 0, 0, 0, 0, 4, tn(8))
  p1$ = str_(tn(1))
   p2$ = str_(tn(3))
    p3$ = str_(tn(5))
     p4$ = str_(tn(7))
      p1$ = combine_para_for_item(p1$, str_(tn(0)))
      p2$ = combine_para_for_item(p2$, str_(tn(2)))
      p3$ = combine_para_for_item(p3$, str_(tn(4)))
      p4$ = combine_para_for_item(p4$, str_(tn(6)))
          tfa = combine_para_for_item(str_(tn(8)), str_(fr))
           fa = time_para(fa, tfa, False, False)
Exit Function
simple_multi_para0_error:
simple_multi_para0 = False 'fa = combine_divide_string(tfa, Trim(str_(tr_p1%)))
       ' Call rational_item(Trim(str_(fr)), it(0), it(1))
End Function


Private Function simple_string_(ByVal s As String, factor1 As String, _
           factor2 As String, cal_float As Boolean) As String
Dim i%, k%, l%, t%, t1%
Dim ty As Byte
Dim ts1(3) As String
Dim ts2(8) As String
Dim ts3(8) As String
Dim tp1(8) As String
Dim tI1(8) As String
Dim tp2(8) As String
Dim tI2(8) As String
Dim S_p_i As para_item_type
'****************************************
On Error GoTo simple_string_error0
ty = string_type(s, "", ts1(0), ts1(1), "")
If ty = 3 Then
 simple_string_ = s
   Exit Function
Else
If InStr(1, s, "F", 0) > 0 Then
 simple_string_ = "F"
  Exit Function
End If
 factor1 = "1"
  factor2 = "1"
If InStr(1, s, ".", 0) > 0 Then
    factor1 = "1"
     factor2 = s
      simple_string_ = s
        Exit Function
ElseIf is_simple(s, 0) Then
 factor1 = "1"
  factor2 = s
   simple_string_ = s
    Exit Function
End If
l% = Len(s)
For i% = 1 To l% - 1
If Mid$(s, i%, 1) >= "A" Then
GoTo simple_string1_mark2
Exit Function
End If
Next i%
For i% = 1 To l%
If Mid$(s, i%, 1) = "&" Then
s = Mid$(s, 1, i% - 1) + "/" + Mid$(s, i% + 1, l% - i%)
'ElseIf Mid$(s, i%, 1) = "#" Then
's = Mid$(s, 1, i% - 1) + "+" + Mid$(s, i% + 1, l% - i%)
End If
Next i%
Call remove_brace(s)
simple_string_ = s
factor1 = "1"
factor2 = s
Exit Function '数
simple_string1_mark2:
t% = string_type(s, ts1(0), ts1(1), ts1(2), ts1(3))
If t% = 0 Then
 If ts1(3) = "" Then
  simple_string_ = s
   factor1 = "1"
    factor2 = s '单项
   Exit Function
 ElseIf ts1(3) <> "" Then '多项
 If read_string(s, S_p_i) = False Then
      simple_string_ = s
       factor1 = s
        factor2 = "1"
         Exit Function
 End If
   Call simple_multi_para_for_sp(S_p_i, ts2(0))
   Call simple_multi_item_for_sp(S_p_i, ts2(1))
      If ts2(1) = "1" Then
       factor2 = ts2(0)
      ElseIf ts2(0) = "1" Then
       factor2 = ts2(1)
      Else
       factor2 = ts2(0) + "*" + ts2(1)
      End If
    factor1 = ""
     simple_string_ = ""
    For i% = 0 To S_p_i.last_it - 1
     ts2(2) = time_string(S_p_i.pA(i%), S_p_i.it(i%), True, False)
      If factor1 = "" Then
          factor1 = ts2(2)
      Else
          If Mid$(ts2(2), 1, 1) = "-" Then
          factor1 = factor1 + ts2(2)
          Else
          factor1 = factor1 + "+" + ts2(2)
          End If
      End If
      If ts2(1) <> "1" Then
       ts2(3) = time_string(S_p_i.pA(i%), _
          time_item(S_p_i.it(0), ts2(1)), False, False)
      If simple_string_ = "" Then
          simple_string_ = ts2(3)
      Else
          simple_string_ = simple_string_ + "+" + ts2(3)
      End If

      Else
       simple_string_ = factor1
      End If
      Next i%
      
  If ts2(0) <> "1" Then
   If ts2(0) = "-" Then
    simple_string_ = "-" + simple_string_
   Else
   simple_string_ = ts2(0) + simple_string_
   End If
  End If
ElseIf t% = 1 Or t% = 2 Then
 simple_string_ = s
  factor1 = s
   factor1 = "1"
ElseIf t% = 3 Then
Call simple_string_(ts1(1), ts2(0), ts2(1), cal_float)
 Call simple_string_(ts1(2), ts2(2), ts2(3), cal_float)
  If ts2(1) <> "1" And ts2(3) <> "1" Then
   Call simple_multi_string(ts2(1), ts2(3), "", "", "", "", _
            "", "", "", 2, ts2(4), True, cal_float)
  End If
    ts1(1) = time_string(ts2(0), ts2(1), True, False)
      ts1(2) = time_string(ts2(2), ts2(3), True, False)
ts1(1) = add_brace(ts1(1), "")
 ts1(2) = add_brace(ts1(2), "root")
  If ts1(2) = "1" Then
   simple_string_ = ts1(1)
  Else
   simple_string_ = ts1(1) + "/" + ts1(2)
  End If
  
End If
End If

l% = Len(s)
For i% = 1 To l% - 1
If Mid$(s, i%, 1) >= "A" Then
simple_string_ = s
Exit Function

End If
Next i%
For i% = 1 To l%
If Mid$(s, i%, 1) = "&" Then
s = Mid$(s, 1, i% - 1) + "/" + Mid$(s, i% + 1, l% - i%)
End If
Next i%
End If
Exit Function
simple_string_error0:
 simple_string_ = "F"
'**************8
End Function

Private Function rational_para(ByVal p$, p1$, p2$) As Boolean
'有理化 1/p$=p2$/p1$ ' p$是根式，p1$是有理式
Dim t As Integer
Dim s(5) As String
Dim tn(5) As Long
Dim ts(1) As String
Dim i%
t = para_type(p$, "", s(0), s(1), s(2))
If t = 3 Then '分数
rational_para = rational_para(s(0), s(3), s(4))
 p2$ = s(4)
  p1$ = time_para(s(3), s(1), False, False)
   Call simple_multi_para(p1$, p2$, "", "", "", "", "", "", "", 2, "", True, False)
ElseIf t = 2 Then '小数
 p2$ = p$
  p1$ = "1"
   rational_para = False
ElseIf t = 1 Then
 
  p1$ = "'" + s(1)
   p2$ = str_(val_(s(0)) * val_(s(1)))
    If Mid$(p2$, 1, 1) = "-" Or Mid$(p2$, 1, 1) = "@" Then
     p1$ = time_para("@1", p1$, True, False)
      p2$ = time_para("@1", p2$, True, False)
    End If
     rational_para = True
ElseIf t = 0 Then '整式
 If s(2) = "" Then
   p2$ = p
    p1$ = "1"
     rational_para = True
 Else
   Call para_type(s(2), "", s(3), s(4), ts(0))
 If ts(0) = "" Then
  Call simple_two_para(s(0), s(3), ts(0))
   If s(0) = "1" Or s(0) = "+1" Or s(0) = "#1" Then
      If s(1) = "1" Then
       s(0) = "1"
      Else
       s(0) = "'" & s(1)
      End If
   ElseIf s(0) = "@1" Or s(0) = "-1" Then
      If s(1) = "@1" Then
       s(0) = "1"
      Else
       s(0) = "@'" & s(1)
      End If
   Else
      If s(1) = "1" Then
       s(0) = s(0)
      Else
       s(0) = s(0) & "'" & s(1)
      End If
   End If
   If s(3) = "1" Or s(3) = "+1" Or s(3) = "#1" Then
      If s(4) = "1" Then
       s(3) = "1"
      Else
       s(3) = "'" & s(4)
      End If
   ElseIf s(3) = "-1" Or s(3) = "@1" Then
      If s(4) = "1" Then
       s(3) = "@1"
      Else
       s(3) = "@'" & s(4)
      End If
   Else
      If s(4) = "1" Then
       s(3) = s(3)
      Else
       s(3) = s(3) & "'" & s(4)
      End If
   End If
p1$ = minus_para(s(0), s(3), False, False)
p2$ = minus_para(time_para(s(0), s(0), False, False), _
          time_para(s(3), s(3), False, False), False, False)
p2$ = time_para(p2$, ts(0), False, False)
If Mid$(p2$, 1, 1) = "-" Or Mid$(p2$, 1, 1) = "@" Then
 p1$ = time_para("-1", p1$, False, False)
 p2$ = time_para("-1", p2$, False, False)
End If
     rational_para = True
    End If
 End If
ElseIf t = 3 Then
 If rational_para(s(1), s(3), p2$) Then
  p1$ = time_para(s(2), s(3), True, False)
    If Mid$(p2$, 1, 1) = "-" Or Mid$(p2$, 1, 1) = "@" Then
     p1$ = time_para("@1", p1$, True, False)
      p2$ = time_para("@1", p2$, True, False)
    End If
   rational_para = True
 Else
  p2$ = str_(val_(value_para(s(0))) / val_(value_para(s(1))))
   p1$ = "1"
    If Mid$(p2$, 1, 1) = "-" Or Mid$(p2$, 1, 1) = "@" Then
     p1$ = time_para("@1", p1$, True, False)
      p2$ = time_para("@1", p2$, True, False)
    End If
   rational_para = False
 End If
End If
End Function
Private Function right_brace(s$, ByVal start%) As Integer
Dim i%, k% ' start% "("后的一
For i% = start% To Len(s$)
 If Mid$(s$, i%, 1) = "(" Then
  k% = k% + 1
 ElseIf Mid$(s$, i%, 1) = ")" Then
  k% = k% - 1
 End If
 If k% = -1 Then
  right_brace = i%
   Exit Function
 End If
Next i%
End Function
Private Function is_right_an_item(s$, ByVal start%) As Boolean
Dim i%, t_rbrace%, t_lbrace% '后的一
Dim ch$
Do
 t_rbrace% = right_brace(s$, start%) '后括号
  If Len(s) = t_rbrace Then '后括号在结尾
     is_right_an_item = True
      Exit Function
  Else '
    t_lbrace% = InStr(t_rbrace% + 1, s$, "(", 0) '下一个前括号
     If t_lbrace% = 0 Then '没有下一个前括号
         t_lbrace% = Len(s$) + 1 '
     End If
     For i% = t_rbrace% + 1 To t_lbrace% - 1 '后括号和下一个前括号之间
      ch$ = Mid$(s$, 1, i%)
       If ch$ = "+" Or ch$ = "-" Or ch$ = "#" Or ch$ = "@" Then '有加减运算
        is_right_an_item = False
         Exit Function
       End If
     Next i%
      start% = t_lbrace% + 1
  End If
Loop Until Len(s) < start%
End Function

Private Function left_brace(s$, start%) As Integer
Dim i%, k% ' start% "("后的一

For i% = start% To 1 Step -1
 If Mid$(s$, i%, 1) = ")" Then
  k% = k% + 1
 ElseIf Mid$(s$, i%, 1) = "(" Then
  k% = k% - 1
 End If
 If k% = -1 Then
  left_brace = i%
   Exit Function
 End If
Next i%
End Function


Public Function abs_string(ByVal s As String) As String
On Error GoTo abs_string_error
If InStr(1, s, "F", 0) > 0 Or s = "" Then
abs_string_error:
abs_string = "F"
Else
If Mid$(s, 1, 1) = "-" Then
abs_string = Mid$(s, 2, Len(s))
Else
abs_string = s
End If
End If
End Function
Public Function sqr_string(ByVal A As String, is_simple As Boolean, cal_float As Boolean) As String
Dim k%, j%, l_b%, sqr_p%
Dim s(3) As String
Dim p(2) As String
Dim f(2) As String
Dim i(1) As String
Dim ty As Byte
Dim S_p As para_item_type
'0整式
'1根式
'2浮点数
'3 分式
On Error GoTo sqr_string_error
If A = "" Or InStr(1, A, "U", 0) > 0 Or InStr(1, A, "V", 0) > 0 Then
 sqr_string = "F"
  Exit Function
ElseIf InStr(1, A, ".", 0) > 0 Then
   cal_float = True
End If
If InStr(1, A, "F", 0) > 0 Or A = "" Then
   sqr_string = "F"
    Exit Function
ElseIf A = "1" Or A = "#1" Or A = "+1" Then
   sqr_string = "1"
   Exit Function
ElseIf A = "0" Then
   sqr_string = "0"
    Exit Function
ElseIf InStr(1, A, "[", 0) > 0 Then
 sqr_string = "F"
  Exit Function
End If
ty = string_type(A, s(0), s(1), s(2), s(3))
If ty = 3 Then
 p(1) = sqr_string(s(2), is_simple, cal_float)
  If InStr(1, p(1), "[", 0) > 0 Or InStr(1, p(1), "'", 0) > 0 Then
   s(1) = time_string(s(1), s(2), False, cal_float)
    p(0) = sqr_string(s(1), False, cal_float)
     sqr_string = divide_string(p(0), s(2), is_simple, cal_float)
  Else
    p(0) = sqr_string(s(1), False, cal_float)
     sqr_string = divide_string(p(0), p(1), is_simple, cal_float)
  End If
    Exit Function
Else
 If s(3) = "" Then
    p(0) = sqr_para(s(1), "", "", cal_float)
    If p(0) = "F" Then
      sqr_string = "F"
       Exit Function
    ElseIf p(0) = "I" Then
     sqr_string = set_squre_root_string_(A)
    Else
     p(1) = sqr_item(s(2), "", "")
      If p(1) = "F" Then
      sqr_string = "F"
       Exit Function
      Else
       sqr_string = time_string(p(0), p(1), True, cal_float)
      End If
    End If
 Else
 Call read_string(A, S_p)
  If S_p.last_it = 2 Then
     sqr_string = set_squre_root_string_(A)
  ElseIf S_p.last_it = 3 Then
    f(0) = sqr_string(S_p.item(0), True, True)
    f(2) = sqr_string(S_p.item(2), True, True)
    If s(0) <> "F" And s(2) <> "F" Then
     f(1) = time_string(f(0), f(2), False, cal_float)
     f(1) = time_string(f(1), "2", False, cal_float)
      i(0) = add_string(S_p.item(1), f(1), True, False)
      i(1) = minus_string(S_p.item(1), f(1), True, False)
      If i(1) = "0" Then
        sqr_string = add_string(f(0), f(2), True, cal_float)
      ElseIf i(0) = "0" Then
        sqr_string = minus_string(f(0), f(2), True, cal_float)
      Else
       sqr_string = set_squre_root_string_(A)
      End If
    Else
     sqr_string = set_squre_root_string_(A)
    End If
  Else
   sqr_string = "F"
  End If
 End If
End If
Exit Function
sqr_string_error:
sqr_string = "F"
End Function

Public Function initial_string(s$) As String
Dim i%, n1%, n2%, n3%, k%, j%
Dim ch1$, ch2$, ch3$
Dim ts$
If s$ = "" Then
 initial_string = "0"
  Exit Function
ElseIf Len(s$) = 1 Then
 initial_string = s$
  Exit Function
End If
Call remove_brace(s$) '去括号
n2% = find_first_item(s$, "+", "-") '分离单项
If n2% = Len(s$) Then
   n2% = 0
End If
'*************************************
If n2% = 0 Then '无加号,单项
'******************************
'确定乘除号
n2% = find_first_item(s$, "*", "/")
  If n2% = Len(s$) Then
    n2% = 0
  End If
'************************************
      If n2% = 0 Then '无乘除
       If Len(s$) = 1 Then '单字母
         initial_string = s$
       Else '多个字母
         n2% = InStr(1, s$, "(", 0) '确定括号
          If n2% > 0 Then '有括号
          '**********************************
            ch1$ = Mid$(s$, 1, n2% - 1) '括号前
           If ch1$ = "" Or ch1$ = "+" Then
              ch1$ = "1"
           ElseIf ch1$ = "-" Then
              ch1$ = "-1"
           Else
              ch1$ = initial_string(ch1$)
           End If
          '************************************
          n1% = right_brace(s$, n2% + 1) '后括号
          '***************************
            ch2$ = Mid$(s$, n2% + 1, n1% - n2% - 1)
            ch2$ = initial_string(ch2$) '括号内
           '*****************************************
            If n1% = Len(s$) Then ''后括号是结尾
            initial_string = time_string(ch1$, ch2$, True, False)
           Else
            If Mid$(s$, n1% + 1, 1) = "^" Then
                ch3$ = Mid$(s$, n1 + 3, Len(s$) - n1% - 2)
            If ch3$ = "" Then
               ch3$ = "1"
            Else
               ch3$ = initial_string(ch3$)
            End If
            If Mid$(s$, n1% + 2, 1) = "2" Then
             ch2$ = time_string(ch2$, ch2$, True, False)
            ElseIf Mid$(s$, n1% + 2, 1) = "3" Then
             initial_string = time_string(ch2$, time_string(ch2$, ch2$, True, False), _
                   True, False)
            End If
           Else
            ch3$ = Mid$(s$, n1 + 1, Len(s$) - n1%)
            If ch3$ = "" Then
               ch3$ = "1"
            Else
               ch3$ = initial_string(ch3$)
            End If
           End If
           initial_string = time_string(ch1$, ch2$, True, False)
            initial_string = time_string(initial_string, _
                  ch3$, True, False)
          End If
         Else '无括号
          ch1$ = Mid$(s$, 1, 1)
           ch3$ = s$
         If ch1$ < "A" Then '分离数
          ch2$ = ""
          '**********************
          Do
           ch2$ = ch2$ & ch1$
            If Len(ch3$) = 1 Then
             ch3$ = ""
             ch1$ = ""
            Else
             ch3$ = Mid$(ch3$, 2, Len(ch3$) - 1)
             ch1$ = Mid$(ch3$, 1, 1)
            End If
           Loop Until ch1$ >= "A" Or ch3$ = ""
           '********************
           If ch2$ = "+" Then
              ch2$ = "1"
           ElseIf ch2$ = "-" Then
              ch2$ = "-1"
           End If
           ch1$ = Mid$(ch3$, 1, 1)
           If ch1$ = "^" Then
            ch1$ = Mid$(ch3$, 1, 2)
             If ch1$ = "2" Then
                ch2$ = time_string(ch2$, ch2$, True, False)
             ElseIf ch1$ = "3" Then
                ch2$ = time_string(ch2$, time_string(ch2$, ch2$, False, False), _
                    True, False)
             End If
             If Len(ch3$) - 2 > 0 Then
               ch3$ = Mid$(ch3$, 3, Len(ch3$) - 2)
             Else
               ch3$ = "1"
             End If
            Else
             If ch3$ = "" Then
               ch3$ = "1"
             End If
            End If
         '***********************************
          initial_string = time_string(ch2$, initial_string(ch3$), True, False)
         Else 'ch1$ < "A" Then '分离数
          ch1 = Mid$(s$, 1, 1)
           ch3$ = Mid$(s$, 2, 1)
             If ch3$ = "^" Then
              ch3$ = Mid$(s$, 3, 1)
                If ch3$ = "2" Then
                 ch1$ = time_string(ch1$, ch1$, True, False)
                ElseIf ch3$ = "3" Then
                 ch1$ = time_string(time_string(ch1$, ch1$, False, False), _
                   ch1$, True, False)
                End If
                If Len(s$) - 3 > 0 Then
                 ch3$ = initial_string(Mid$(s$, 4, Len(s$) - 3))
                Else
                  ch3$ = "1"
                End If
            Else
              If Len(s$) - 1 > 0 Then
               ch3$ = initial_string(Mid$(s$, 2, Len(s$) - 1))
              Else
               ch3$ = "1"
              End If
            End If
             initial_string = time_string(ch1$, ch3$, True, False)
         End If 'ch1$ < "A" Then '分离数
         End If
        End If '多个字母
    Else '无乘除
    ch1$ = Mid$(s$, n2% + 1, 1)
     If ch1$ = "*" Then
         initial_string = time_string(initial_string(Mid$(s$, 1, n2%)), _
              initial_string(Mid$(s$, n2% + 2, Len(s$) - n2% - 1)), True, False)
     ElseIf ch1$ = "/" Then
         initial_string = divide_string(initial_string(Mid$(s$, 1, n2%)), _
              initial_string(Mid$(s$, n2% + 2, Len(s$) - n2% - 1)), True, False)
     ElseIf ch1$ = "(" Then
       n1% = right_brace(s$, n2% + 2)
         ch1$ = Mid$(s$, 1, n2%)
          If ch1$ = "" Then
            ch1$ = "1"
          Else
           ch1$ = initial_string(ch1$)
          End If
          If Len(s$) > n1% Then
          ch3$ = Mid$(s$, n1% + 1, 1)
          Else
          ch3$ = ""
          End If
          initial_string = initial_string(Mid$(s$, n2% + 2, n1% - n2% - 2))
          If ch3$ = "^" Then
           ch2$ = Mid$(s$, n1% + 2, 1)
                        If ch2$ = "2" Then
              initial_string = time_string(initial_string, _
               initial_string, True, False)
            ElseIf ch2$ = "3" Then
              initial_string = time_string(time_string(initial_string, _
               initial_string, False, False), _
               initial_string, True, False)
            End If
             ch3$ = Mid$(s$, n1% + 3, Len(s$) - n3% - 2)
          ElseIf ch3$ = "*" Then
           ch3$ = Mid$(s$, n1% + 2, Len(s$) - n1% - 1)
          ElseIf ch3$ = "/" Then
           ch3$ = Mid$(s$, n1% + 2, Len(s$) - n1% - 1)
            initial_string = time_string(ch1$, _
                initial_string, True, False)
            initial_string = divide_string(initial_string, _
                ch3$, True, False)
           Exit Function
          Else
             ch3$ = Mid$(s$, n1% + 1, Len(s$) - n1%)
          End If
             If ch3$ = "" Then
              ch3$ = "1"
             End If
         initial_string = time_string(ch1$, _
                initial_string, True, False)
         initial_string = time_string(initial_string, _
              initial_string(ch3$), True, False)
     ElseIf ch1$ = "'" Then
       If Mid$(s$, n2% + 1, 1) = "(" Then
         n1% = right_brace(s$, n2% + 2)
       Else
         ch1$ = Mid$(s$, n2% + 1, 1)
          ch3$ = Mid$(s$, n2% + 1, Len(s$) - n2%)
         If ch1$ < "A" Then '分离数
            ch2$ = ""
         Do
          ch2$ = ch2$ & ch1$
           If Len(ch3$) = 1 Then
           ch3$ = ""
           ch1$ = ""
           Else
           ch3$ = Mid$(ch3$, 2, Len(ch3$) - 1)
           ch1$ = Mid$(ch3$, 1, 1)
           End If
         Loop Until ch1$ >= "A" Or ch3$ = ""
         If ch2$ = "+" Then
            ch2$ = "1"
         ElseIf ch2$ = "-" Then
            ch2$ = "-1"
         End If
         If ch3$ = "" Then
          ch3$ = "1"
         End If
          initial_string = Mid$(s$, 1, n2% - 1)
          initial_string = time_string("'" & ch2$, initial_string, True, False)
          initial_string = time_string(initial_string(ch3$), initial_string, True, False)
        Else
         initial_string = Mid$(s$, 1, n2% - 1)
         initial_string = time_string(initial_string, Mid$(s$, n2%, 2), True, False)
         If Len(s$) > n2% + 2 Then
          initial_string = time_string(initial_string, Mid$(s$, n2% + 3, Len(s$) - n2% - 2), True, False)
         End If
       End If
    End If
   End If
   End If
  Else
   initial_string = add_string(initial_string(Mid$(s$, 1, n2%)), _
         initial_string(Mid$(s$, n2% + 1, Len(s$) - n2%)), True, False)
 End If
'End If
End Function
Private Function sqr_para(ByVal p1$, p2$, p3$, cal_float As Boolean) As String
Dim s(8) As String
Dim ty As Byte
'Dim t As Byte
Dim n(2) As Long
Dim sig As String
'0整式
'1根式
'2浮点数
'3 分式
On Error GoTo sqr_para_erro
'*********************'
'去括号
Call remove_brace(p1$)
'*****************
'特殊值
If InStr(1, p1$, "F", 0) > 0 Or InStr(1, p1$, "[", 0) > 0 Then '重复开方
 sqr_para = "F"
  Exit Function
ElseIf p1$ = "1" Or p1$ = "+1" Or p1$ = "#1" Then '单位开方
 p2$ = "1"
  p3$ = "1"
   sqr_para = "1"
  Exit Function
ElseIf p1$ = "0" Then '单位开方
 p2$ = "0"
  p3$ = "1"
   sqr_para = "0"
  Exit Function
End If
'*************************
ty = para_type(p1$, s(0), s(1), s(2), s(3)) '分解
If ty = 3 Then '分式
   s(3) = sqr_para(s(1), s(4), s(5), cal_float) '分母开方
    s(6) = sqr_para(s(2), s(7), s(8), cal_float) '分子开方
   If s(3) = "F" Or s(6) = "F" Then '失败
     sqr_para = "F"
   Else
    If s(8) = "1" Then '分母无根号
      sqr_para = divide_para(s(3), s(7), True, cal_float)
    Else '有理化
      sqr_para = divide_para(time_para(s(3), "'" & s(8), False, cal_float), _
           time_string(s(7), s(8), False, cal_float), True, cal_float)
    End If
   End If
Else
 If s(3) = "" Then '单项
   If s(2) = "1" Then
    n(0) = val(s(1))
     If n(0) > 0 Then
      Call sq_root(n(0), n(1), n(2)) '分解整树
       If n(2) = 1 Then
        p2 = Trim(str(n(1)))
        p3 = "1"
        sqr_para = p2 'Trim(Str(N(1)))
       Else
        If n(1) = 1 Then
          p2 = "1"
          p3 = Trim(str(n(2)))
         sqr_para = "'" & Trim(str(n(2)))
        Else
          p2 = Trim(str(n(1)))
          p3 = Trim(str(n(2)))
         sqr_para = p2 & "'" & p3
        End If
       End If
     Else
       sqr_para = "F" '重复开方
     End If
    Else
     sqr_para = "I"
    End If
 Else
     Call para_type(s(3), s(4), s(5), s(6), s(7)) '再分解
   If s(7) = "" Then '两项
     If s(2) <> "1" And s(6) <> "1" Then '
        sqr_para = "I"
         Exit Function
     ElseIf s(2) = "1" And s(6) <> "1" Then
       If val(s(1)) < 0 Then
        sqr_para = "F"
       Else
        sig = Mid$(s(5), 1, 1)
        If sig = "-" Or sig = "@" Then
           sig = "-"
        ElseIf sig = "+" Or sig = "#" Then
           sig = "+"
        Else
           sig = "+"
        End If
        s(3) = time_string("-1", s(1), True, False)
        s(1) = divide_string(s(5), "2", False, False)
        s(7) = time_para(s(5), s(5), False, False)
        s(7) = time_para(s(7), s(6), False, False)
       End If
     ElseIf s(2) <> "1" And s(6) = "1" Then
        If val(s(5)) < 0 Then
         sqr_para = "F"
        Else
         sig = Mid$(s(1), 1, 1)
         If sig = "-" Or sig = "@" Then
           sig = "-"
         ElseIf sig = "+" Or sig = "#" Then
           sig = "+"
         Else
           sig = "+"
         End If
        s(3) = time_string("-1", s(5), True, False)
        s(1) = divide_para(s(1), "2", False, False)
        s(7) = time_para(s(1), s(1), False, False)
        s(7) = time_para(s(7), s(2), False, False)
        End If
    End If
        sqr_para = sqr_para_for_two_root(s(3), s(7), sig)
   Else
    sqr_para = "I"
   End If
  End If
End If
Exit Function
sqr_para_erro:
sqr_para = "F"
End Function

Private Function sqr_item(ByVal s1$, S2$, s3$) As String
Dim i%
Dim ts(1) As String
Dim tp(1) As String
On Error GoTo sqr_item_error
If Len(s1$) = 1 Then
 If Asc(s1$) < 0 Or InStr(1, s1$, "F", 0) > 0 Then
  sqr_item = "F"
   Exit Function
 End If
ElseIf InStr(1, s1$, "U", 0) > 0 Or InStr(1, s1$, "V", 0) > 0 Then
   sqr_item = "F"
   Exit Function
End If
i% = InStr(1, s1$, "/", 0) ' 分式
If i% > 0 Then
 ts(0) = Mid$(s1$, 1, i% - 1) '分子
  ts(1) = Mid$(s1$, i% + 1, Len(s1$) - i%) '分母
  tp(0) = ""
   tp(1) = ""
Do '分出双因子
If Mid$(ts(1), 1, 1) = Mid$(ts(1), 2, 1) Then
    tp(0) = tp(0) + Mid$(ts(1), 1, 1) '第一因子
     ts(1) = Mid$(ts(1), 3, Len(ts(1)) - 2)
Else
    tp(1) = tp(1) + Mid$(ts(1), 1, 1) '第二因子
     ts(1) = Mid$(ts(1), 2, Len(ts(1)) - 1)
End If
Loop Until Len(ts(1)) = 0
'********************************************'
 s3$ = time_string(tp(0), tp(1), True, False)
  ts(0) = time_string(tp(1), ts(0), True, False)
   tp(0) = ""
    tp(1) = ""
Do
If Len(ts(0)) > 1 Then
If Mid$(ts(0), 1, 1) = Mid$(ts(0), 2, 1) Then
    tp(0) = tp(0) + Mid$(ts(0), 1, 1)
     ts(0) = Mid$(ts(0), 3, Len(ts(0)) - 2)
Else
    tp(0) = tp(0) + Mid$(ts(0), 1, 1)
     ts(0) = Mid$(ts(0), 2, Len(ts(0)) - 1)
End If
ElseIf Len(ts(0)) > 0 Then
    tp(0) = tp(0) + Mid$(ts(0), 1, 1)
     ts(0) = Mid$(ts(0), 2, Len(ts(0)) - 1)
End If

Loop Until Len(ts(0)) = 0
Call set_squre_root_string(tp(1), tp(1))
If InStr(1, tp(1), "F", 0) > 0 Then
 sqr_item = "F"
Else
sqr_item = tp(0) + tp(1) + "/" + s3$
S2$ = tp(0) + "/" + s3$
s3$ = tp(1)
End If
Else

  tp(0) = ""
   tp(1) = ""
If InStr(1, s1$, "+", 0) > 1 Or InStr(1, s1$, "-", 0) > 1 Then
 tp(0) = "1"
 Call set_squre_root_string(s1$, tp(1))
 sqr_item = tp(1)
  Exit Function
Else
Do
If Len(s1$) >= 2 Then
If Mid$(s1$, 1, 1) = Mid$(s1$, 2, 1) Then
    tp(0) = tp(0) + Mid$(s1$, 1, 1)
     s1$ = Mid$(s1$, 3, Len(s1$) - 2)
Else
    tp(1) = tp(1) + Mid$(s1$, 1, 1)
     s1$ = Mid$(s1$, 2, Len(s1$) - 1)
End If
ElseIf Len(s1$) = 1 Then
     tp(1) = tp(1) + Mid$(s1$, 1, 1)
     s1$ = ""
End If
Loop Until Len(s1$) = 0
End If
If tp(1) = "" Then
 sqr_item = tp(0)
Else
 Call set_squre_root_string(tp(1), tp(1))
 If InStr(1, tp(1), "F", 0) > 0 Then
 sqr_item = "F"
 Else
 If tp(0) = "1" Then
 sqr_item = tp(1)
 Else
 sqr_item = tp(0) + tp(1)
 End If
 End If
End If
If tp(0) = "" Then
S2$ = "1"
Else
S2$ = tp(0)
End If
If tp(1) = "" Then
s3$ = "1"
Else
If Len(tp(1)) = 1 And Asc(tp(1)) < 0 Then
 If S2$ = "1" Then
  S2$ = tp(1)
 Else
 S2$ = S2$ + tp(1)
 End If
 s3$ = "1"
Else
 s3$ = tp(1)
End If
End If

End If
Exit Function
sqr_item_error:
sqr_item = "F"
End Function

Public Function is_less_than(ByVal s1$, ByVal S2$) As Boolean
s1$ = value_para(s1$)
 S2$ = value_para(S2$)
If s1$ = "F" Or S2$ = "F" Then
  is_less_than = False
Else
 If val_(s1$) < val_(S2$) Then
 is_less_than = True
 End If
End If
End Function
Private Sub read_para(ByVal p As String, items As para_item_type)
Dim tp As String
Dim tp1 As String
Dim i%, j%
p = unsimple_para(p)
items.last_it = 0
tp = p
If tp = "" Then
Exit Sub
End If
Do
     ReDim Preserve items.pA(items.last_it) As String
      ReDim Preserve items.it(items.last_it) As String
Call para_type(tp, "", items.pA(items.last_it), _
       items.it(items.last_it), tp)
       items.last_it = items.last_it + 1
Loop Until tp = ""
End Sub


Private Function time_item(ByVal i1 As String, ByVal i2 As String) As String
Dim i%, j%, n1%, n2%, tn%, tn2%
Dim ch(1 To 6) As String
Dim temp_I As String
On Error GoTo time_item_error
If InStr(1, i1, "F", 0) > 0 Or InStr(1, i2, "F", 0) > 0 Then
 time_item = "F"
  Exit Function
End If
temp_I = i1
If i1 = "1" Then
time_item = i2
Exit Function
ElseIf i2 = "1" Then
time_item = i1
Exit Function
End If
Call remove_brace(i1)
Call remove_brace(i2)
n1% = InStr(1, i1$, "/", 0)
n2% = InStr(1, i2$, "/", 0)
If n1% = 0 And n2% = 0 Then '单项
 Call read_sqr_from_item(i1, 1, ch(1), ch(3), ch(5))
 Call read_sqr_from_item(i2, 1, ch(2), ch(4), ch(6))
 If ch(3) <> "1" And ch(4) <> "1" Then
  If i1 = i2 Then
   time_item = time_string(ch(1), ch(2), False, False)
    If InStr(1, ch(3), "+", 0) = 0 And InStr(1, ch(3), "-", 0) = 0 And _
         InStr(1, ch(3), "#", 0) = 0 And InStr(1, ch(3), "@", 0) = 0 Then
          If time_item = "1" Or time_item = "+1" Or time_item = "#1" Then
            time_item = ch(3)
          ElseIf time_item = "-1" Or time_item = "@1" Then
            time_item = "-" + ch(3)
          Else
          time_item = time_item(time_item, ch(3))
          End If
    Else
          time_item = time_string(time_item, ch(3), True, False)
    End If
     Exit Function
  Else
    time_item = sqr_string(time_string(ch(3), ch(4), False, False), True, False)
    If (InStr(1, time_item, "+", 0) = 0 And InStr(1, time_item, "-", 0) = 0 And _
         InStr(1, time_item, "#", 0) = 0 And InStr(1, time_item, "@", 0) = 0) Or _
            InStr(1, time_item, "[", 0) > 0 Then
             ch(6) = time_item(ch(1), ch(2))
              If ch(6) = "1" Or ch(6) = "+1" Or ch(6) = "#1" Then
               time_item = time_item
              ElseIf ch(6) = "-1" Or ch(6) = "@1" Then
               time_item = "-" + time_item
              Else
               time_item = time_item(ch(6), time_item)
              End If
    Else
     time_item = time_string(time_item(ch(1), ch(2)), time_item, True, False)
    End If
   Exit Function
  End If
 ElseIf ch(3) <> "1" Then
    i2 = time_item(ch(1), i2)
     If ch(5) <> "[0]" Then
       If i2 < ch(5) Then
        time_item = i2 & ch(5) 'time_string(display_string_(squre_root_string(tn), 0), i2, True, False)
       Else
        time_item = ch(5) & i2
       End If
     Else
      If ch(3) <> "1" Then
       ch(3) = "'" & ch(3)
       If i2 < ch(2) Or Mid$(ch(2), 1, 1) = "'" Then
          time_item = i2 & ch(3)
       Else
          time_item = ch(3) & i2
       End If
      Else
       time_item = i2
      End If
     End If
   Exit Function
 ElseIf ch(4) <> "1" Then '
     i1 = time_item(ch(2), i1)
       If ch(6) <> "[0]" Then
         If i1 < ch(6) Then
          time_item = i1 & ch(6) 'time_item = i2 + i1 'time_string(display_string_(squre_root_string(tn), 0), i1, True, False)
         Else
          time_item = ch(6) & i1
         End If
       Else
        If ch(4) <> "1" Then
         ch(4) = "'" & ch(4)
          If i1 < ch(4) Or Mid$(ch(4), 1, 1) = "'" Then
           time_item = i1 & ch(4)
          Else
           time_item = ch(4) & i1
          End If
        Else
         time_item = i1
        End If
       End If
   Exit Function
 Else
 If i1 < i2 Then
  temp_I = i1 & i2
 Else
  temp_I = i2 & i1
 End If
  For i% = Len(temp_I) - 1 To 1 Step -1 '排序
   For j% = 1 To i%
    ch(1) = Mid$(temp_I, j%, 1)
     ch(2) = Mid$(temp_I, j% + 1, 1)
  If (ch(1) > ch(2)) And (ch(1) <> empty_char And ch(2) <> empty_char) Then
   If j% > 1 Then
   i1 = Mid$(temp_I, 1, j% - 1)
   Else
    i1 = ""
   End If
    If i1 < ch(2) And ch(2) < ch(1) Then
     i1 = i1 & ch(2) & ch(1)
    ElseIf i1 < ch(1) And ch(1) < ch(2) Then
     i1 = i1 & ch(1) & ch(2)
    ElseIf ch(1) < ch(2) And ch(2) < i1 Then
     i1 = ch(1) & ch(2) & i1
    ElseIf ch(1) < i1 And i1 < ch(2) Then
     i1 = ch(1) & i1 & ch(2)
    ElseIf ch(2) < ch(1) And ch(1) < i1 Then
     i1 = ch(2) & ch(1) & i1
    ElseIf ch(2) < i1 And i1 < ch(1) Then
     i1 = ch(2) & i1 & ch(1)
    End If
   If Len(temp_I) > j% + 1 Then
    i1 = i1 + Mid$(temp_I, j% + 2, Len(temp_I) - j% - 1)
   End If
   temp_I = i1
  End If
  Next j%
  Next i%
time_item = simple_item_for_squre_root(temp_I)
End If
End If
Exit Function
time_item_error:
time_item = "F"
End Function

Public Function combine_item_with_para(ByVal pA As String, _
     ByVal i As String, is_simple As Boolean) As String
Dim T_V$, v1$, v2$, v3$ '
Dim ch$
Dim t_i As String
Dim k%, m%, l%, b1%, b2%, st%, la%
Dim tn(1) As Integer
Dim tm(1) As Integer
If InStr(1, pA, "F", 0) > 0 Or InStr(1, i, "F", 0) > 0 Or pA = "" Or i = "" Then
 combine_item_with_para = "F"
  Exit Function
End If
On Error GoTo combine_item_with_para_error
'If Len(i) = 1 And Asc(i) < 0 Then
'm% = from_char_to_no(i)
'combine_item_with_para = time_string(pa, squre_root_string(m%), True, False)
'Exit Function
'Else
If pA = "0" Or pA = "" Then
 combine_item_with_para = "0"
  Exit Function
End If
If is_simple = True Then
 pA = simple_para(pA, "", "")
End If
If i = "1" Then
 If Mid$(pA, 1, 1) = "@" Then
  combine_item_with_para = "-" & Mid$(pA, 2, Len(pA) - 1)

 Else
 combine_item_with_para = pA
 End If
  Exit Function
ElseIf pA = "1" Or pA = "+1" Or pA = "#1" Then
 combine_item_with_para = i
  Exit Function
ElseIf pA = "-1" Or pA = "@1" Then
 combine_item_with_para = "-" + i
  Exit Function
End If
tn(0) = InStr(2, pA, "#", 0)
If tn(0) = 0 Then
tn(0) = InStr(2, pA, "+", 0)
If tn(0) = 0 Then
tn(0) = InStr(2, pA, "@", 0)
If tn(0) = 0 Then
tn(0) = InStr(2, pA, "-", 0)
End If
End If
End If
tn(1) = InStr(2, i, "#", 0)
If tn(1) = 0 Then
tn(1) = InStr(2, i, "+", 0)
If tn(1) = 0 Then
tn(1) = InStr(2, i, "@", 0)
If tn(1) = 0 Then
tn(1) = InStr(2, i, "-", 0)
End If
End If
End If
If tn(0) = 0 And tn(1) = 0 Then
 combine_item_with_para = pA + i
 Exit Function
End If
If tn(1) > 0 Or Mid$(i, 1, 1) = "'" Then
 combine_item_with_para = time_string(pA, i, is_simple, False)
ElseIf InStr(1, i, "[", 0) > 0 Then 'Asc(i) < 0 And i <> "\" Then
 If Len(i) = 1 Then
  If InStr(1, pA, "'", 0) = 0 Then
   If pA = "1" Then
    combine_item_with_para = i
   ElseIf pA = "-1" Or pA = "@1" Then
    combine_item_with_para = "-" + i
   Else
    If i = "1" Then
     combine_item_with_para = pA
    Else
     combine_item_with_para = pA + i
    End If
   End If
  Else
   combine_item_with_para = sqr_string(time_string(time_string(pA, pA, False, False), _
            read_sqr_from_string(i, 0, v3$), False, False), True, False)
   If val_(value_string(pA)) < 0 Then
    T_V$ = combine_item_with_para
    combine_item_with_para = ""
    st% = 1
    Do
    b1 = InStr(st%, T_V$, "(", 0)
    If b1% = 0 Then
    b1% = Len(T_V$) + 1
    b2% = Len(T_V$) + 1
    Else
    b2 = right_brace(T_V$, b1% + 1)
    End If
    For l% = st% To b1% - 1
     ch$ = Mid$(T_V$, l%, 1)
     If ch = "+" Then
      combine_item_with_para = combine_item_with_para + "-"
     ElseIf ch = "-" Then
      If l% > 1 Then
      combine_item_with_para = combine_item_with_para + "+"
      End If
     ElseIf ch$ = "#" Then
      combine_item_with_para = combine_item_with_para + "@"
     ElseIf ch$ = "@" Then
      If l% > 1 Then
      combine_item_with_para = combine_item_with_para + "#"
      End If
     Else
      combine_item_with_para = combine_item_with_para + ch
      If l% = 1 Then
       combine_item_with_para = "-" + combine_item_with_para
      End If
     End If
    Next l%
    If b2% > b1% Then
     combine_item_with_para = combine_item_with_para + _
        Mid$(T_V$, b1%, b2% - b1% + 1)
    End If
    st% = b2% + 1
    Loop While st% <= Len(T_V$)
    'combine_item_with_para = time_string("-1", combine_item_with_para, True, False)
   End If
  End If
 Else
  st% = InStr(1, i, "[", 0)
  If st% > 0 Then
   t_i = Mid$(i, 1, st% - 1)
   v1$ = read_sqr_no_from_string(i, st%, la%, "")
    st% = InStr(la% + 1, i, "[", 0)
     If st% > 0 Then
      t_i = t_i + Mid$(i, la% + 1, la% - st% - 1)
      v2$ = read_sqr_no_from_string(i, st%, st%, "")
     Else
      v2$ = "1"
       If Len(i) > la% Then
       t_i = t_i = t_i + Mid$(i, la% + 1, Len(i) - la%)
       End If
     End If
   Else
    v1$ = "1"
    v2$ = "1"
    t_i = i
   End If
   If t_i = "" Then
    t_i = "1"
   End If
  If v1$ = v2$ And v1$ <> "1" Then 'Mid$(i, 2, 1) = Mid$(i, 1, 1) Then
   combine_item_with_para = time_string(time_string(pA, v1$, False, False), t_i, True, False)
   'If Len(i) > 2 Then
    'combine_item_with_para = time_string(combine_item_with_para, _
          Mid$(i, 3, Len(i) - 2), False, False)
   'End If
  ElseIf v1$ <> "1" And v2$ <> "1" Then
   v1$ = time_string(v1$, v2$, False, False)
    v1$ = sqr_string(v1$, False, False)
     pA = combine_item_with_para(pA, t_i, is_simple)
      combine_item_with_para = pA + v1$ 'time_string(time_string(pa, v1$, False, False), t_i, True, False)
  ElseIf v1$ <> "1" Then
   If t_i <> "1" Then
    v1$ = sqr_string(v1$, False, False)
     pA = combine_item_with_para(pA, t_i, is_simple)
      combine_item_with_para = pA + v1$ 'time_string(time_string(pa, v1$, False, False), t_i, True, False)
   Else
    If tn(0) = 0 Then
     If pA = "1" Then
     combine_item_with_para = i
     Else
     combine_item_with_para = pA + i
     End If
    Else
     If Mid$(pA, 1, 1) = "(" And Mid$(pA, Len(pA), 1) = ")" Then
     combine_item_with_para = pA + i
     Else
     combine_item_with_para = "(" + pA + ")" + i
     End If
    End If
   End If
  Else
   If pA = "1" Then
    combine_item_with_para = i
   ElseIf pA = "-1" Or pA = "@1" Then
    combine_item_with_para = "-" + i
   Else
    If i = "1" Then
     combine_item_with_para = pA
    Else
     combine_item_with_para = pA + i
    End If
   End If
   'combine_item_with_para = pa + i
  End If
 End If
Else
 k% = InStr(1, pA, "(", 0)
If k% = 0 Then
If i <> "1" Then
 pA = add_brace(pA, "para")
 End If
Else
If Mid$(pA, 1, 1) = "@" And (InStr(1, pA, "(", 0) > 0 Or _
    (InStr(2, pA, "#", 0) = 0 And InStr(1, pA, "@", 0) = 0)) Then
 pA = "-" + Mid$(pA, 2, Len(pA) - 1)
End If
End If
If InStr(1, pA, "(", 0) = 0 Then
 pA = add_brace(pA, "para")
End If
i = add_brace(i, "para")
If Mid$(pA, 1, 1) = "@" Then
 pA = "-" + Mid$(pA, 2, Len(pA) - 1)
End If
If i = "" Or i = "1" Then
 combine_item_with_para = pA
ElseIf pA = "1" Then
 combine_item_with_para = i
ElseIf pA = "-1" Or pA = "@1" Then
 combine_item_with_para = "-" + i
ElseIf pA = "0" Then
 combine_item_with_para = ""
Else
combine_item_with_para = pA + i
End If
End If
Exit Function
combine_item_with_para_error:
 combine_item_with_para = "F"
End Function

Private Function simple_para(ByVal pA As String, pA1 As String, factor As String) As String
Dim k%, i%, t%
Dim s(1) As String
Dim tn(1) As Long
Dim tp As para_item_type
On Error GoTo simple_para_error
If InStr(1, pA, "F", 0) > 0 Then
 simple_para = "F"
  Exit Function
ElseIf pA = "#0" Or pA = "0" Or pA = "@0" Or pA = "+0" Or pA = "-0" Then
 simple_para = "0"
  Exit Function
End If
'******************************************************************
If InStr(1, pA, ".", 0) > 0 Then
 simple_para = pA
  pA1 = pA
   factor = "1"
    Exit Function
ElseIf is_simple(pA, 1) Then
 simple_para = pA
  pA1 = pA
   factor = "1"
    Exit Function
End If
'*****************************************
k% = InStr(1, pA, "/", 0)
If t% > 0 Then
factor = "1"
 pA1 = pA
  simple_para = pA
   Exit Function
ElseIf k% = 0 Then
  Call read_para(pA, tp)
   If tp.last_it < 2 Then
    GoTo simple_para_mark0
   End If
tn(0) = val_(tp.pA(0))
 tn(1) = val_(tp.it(0))
For i% = 1 To tp.last_it - 1
tn(0) = l_gcd(val_(tp.pA(i%)), tn(0))
tn(1) = l_gcd(val_(tp.it(i%)), tn(1))
Next i%
For i% = 0 To tp.last_it - 1
 tp.pA(i%) = str_(val_(tp.pA(i%)) / tn(0))
  tp.it(i%) = str_(val_(tp.it(i%)) / tn(1))
Next i%
'***********************************************************
 pA1 = para_from_para_item(tp)
  factor = combine_para_for_item(str_(tn(0)), str_(tn(1)))
   If factor <> "1" And factor <> "@1" Then
    simple_para = factor + add_brace(pA1, "para")
   ElseIf factor = "@1" Then
    simple_para = "-" + add_brace(pA1, "para")
   Else
    simple_para = pA1
   End If
   Exit Function
End If
simple_para_mark0:
i% = InStr(1, pA, ".", 0)
 If i% > 0 And i% + 4 <= Len(pA) Then
   pA = Mid$(pA, 1, i% + 4)
  If i% = 1 Then
   pA = "0" + pA
  End If
  End If
 factor = pA
 simple_para = pA
  pA1 = "1"
 Exit Function
simple_para_error:
simple_para = "F"
End Function

Public Function remove_brace(p As String) As Boolean
Dim i%
Dim k%
If p = "" Then
 Exit Function
End If
If Mid$(p, 1, 1) = "(" And Mid$(p, Len(p), 1) = ")" Then
For i% = 1 To Len(p)
If Mid$(p, i%, 1) = "(" Then
 k% = k% + 1
ElseIf Mid$(p, i%, 1) = ")" Then
 k% = k% - 1
End If
 If k% = 0 And i% < Len(p) Then
   GoTo remove_brace_mark0
 End If
Next i%
p = Mid$(p, 1, Len(p) - 1)
 p = Mid$(p, 2, Len(p) - 1)
  remove_brace = True
End If
remove_brace_mark0:
End Function

Public Function unsimple_para(ByVal p As String) As String
Dim n%, m%, k%, i%, b_1%
Dim t_p(3) As String
Dim p_i As para_item_type
Call remove_brace(p)
If InStr(1, p, "&", 0) > 0 Then
 unsimple_para = p
 Exit Function
End If
If Mid$(p, Len(p), 1) = ")" Then
 b_1% = left_brace(p, Len(p) - 1)
End If
k% = InStr(1, p, "(", 0)
 m% = InStr(2, p, "#", 0)
  n% = InStr(2, p, "@", 0)
   If (n% > 0 And n% < m%) Or m% = 0 Then
    m% = n%
   End If
If m% < b_1% And m% > 1 Then
unsimple_para = Mid$(p, 1, m% - 1) + unsimple_para(Mid$(p, m%, Len(p) - m% + 1))
Exit Function
End If
n% = InStr(1, p, "'", 0)
If k% = 0 Then
 unsimple_para = p
  Exit Function
End If
If k% > 0 And (k% < m% Or m% = 0) Then 'And _
     (k% < n% Or n% = 0) Then
If right_brace(p, k% + 1) = Len(p) Then
t_p(0) = Mid$(p, 1, k% - 1)
 p = Mid$(p, k% + 1, Len(p) - k% - 1)
  If t_p(0) = "-" Or t_p(0) = "@" Then
   t_p(0) = "@1"
  ElseIf t_p(0) = "" Then
   t_p(0) = "1"
  End If
Call read_para(p, p_i)
If n% < k% And n% > 0 Then
If n% = 1 Then
t_p(1) = "1"
Else
t_p(1) = Mid$(t_p(0), 1, n% - 1)
If t_p(1) = "#" Then
    t_p(1) = "#1"
ElseIf t_p(1) = "@" Then
    t_p(1) = "@1"
End If
End If
t_p(2) = Mid$(t_p(0), n% + 1, Len(t_p(0)) - n%)
Else
t_p(1) = t_p(0)
t_p(2) = "1"
End If
'Call para_type(t_p(0), t_p(1), t_p(2), "")
For i% = 0 To p_i.last_it - 1
 Call sqr_para(str_(val_(p_i.it(i%)) * val_(t_p(2))), t_p(3), p_i.it(i%), False)
  p_i.pA(i%) = str_((val_(p_i.pA(i%)) * val_(t_p(3))) * val_(t_p(1)))
Next i%
  unsimple_para = para_from_para_item(p_i)
Else
unsimple_para = p
End If
Else
 unsimple_para = p
End If

End Function

Private Function read_string(ByVal s As String, S_p_i As para_item_type) As Boolean
Dim ts As String
On Error GoTo read_string_error
If s = "" Then
 Exit Function
End If
ts = s
Do While ts <> ""
ReDim Preserve S_p_i.pA(S_p_i.last_it) As String
ReDim Preserve S_p_i.it(S_p_i.last_it) As String
ReDim Preserve S_p_i.item(S_p_i.last_it) As String
Call string_type(ts, S_p_i.item(S_p_i.last_it), S_p_i.pA(S_p_i.last_it), _
      S_p_i.it(S_p_i.last_it), ts)
S_p_i.last_it = S_p_i.last_it + 1
Loop
read_string = True
Exit Function
read_string_error:
read_string = False
End Function
Public Function set_string(ByVal s$) As String '输入表达式设置成标准形
Dim k%, k1%, k2%, k3%, k4%, m%
Dim ts(2) As String
k1% = InStr(2, s$, "+", 0)
 m% = InStr(2, s$, "#", 0)
If m% < k1% And m% > 0 Then
 k1% = m%
End If
k2% = InStr(2, s$, "-", 0)
 m% = InStr(2, s$, "@", 0)
If m% < k2% And m% > 0 Then
 k2% = m%
End If
If k1% = 0 Or (k2% < k1% And k2 > 0) Then
 k% = k2%
Else
 k% = k1%
End If
'**************************************
If k% = 0 Then
  m% = InStr(2, s$, "/", 0)
   If m% = 0 Then
       m% = InStr(1, s$, "*", 0)
        If m% = 0 Then
         set_string = s$
        Else
         ts(0) = Mid$(s$, 1, m% - 1)
          ts(1) = Mid$(s$, m% + 1, Len(s$) - m%)
           set_string = time_string(ts(0), set_string(ts(1)), True, False)
        End If
   Else
    ts(0) = Mid$(s$, 1, m% - 1)
    ts(1) = Mid$(s$, m% + 1, Len(s$) - m%)
    set_string = divide_string(set_string(ts(0)), set_string(ts(1)), True, False)
   End If
Else
  k3% = InStr(2, s$, "(", 0)
   If k3% = 0 Then
     ts(0) = Mid$(s$, 1, k% - 1)
     ts(1) = Mid$(s$, k%, Len(s$) - k% + 1)
      set_string = add_string(set_string(ts(0)), set_string(ts(1)), True, False)
   Else
   If k% < k3% Then
    ts(0) = Mid$(s$, 1, k% - 1)
     ts(1) = Mid$(s$, k%, Len(s$) - k% + 1)
      set_string = add_string(set_string(ts(0)), set_string(ts(1)), True, False)
   Else
    k4% = right_brace(s$, k3% + 1)
    If k% > k4% Then
     ts(0) = Mid$(s$, 1, k% - 1)
      ts(1) = Mid$(s$, k%, Len(s$) - k% + 1)
       set_string = add_string(set_string(ts(0)), set_string(ts(1)), True, False)
    Else
           m% = InStr(k4%, s$, "/", 0)
      If m% > 0 Then
       ts(0) = Mid$(s$, 1, m% - 1)
        ts(1) = Mid$(s$, m% + 1, Len(s$) - m%)
         set_string = divide_string(set_string(ts(0)), set_string(ts(1)), True, False)
      Else
       ts(0) = Mid$(s$, 1, k3% - 1)
        ts(2) = Mid$(s$, k4% + 1, Len(s$) - k4%)
       ts(1) = Mid$(s$, k3% + 1, k4% - k3% - 1)
        If Mid$(ts(0), Len(ts(0)), 1) = "*" Then
           ts(0) = Mid$(ts(0), 1, Len(ts(0)) - 1)
        End If
        If Mid$(ts(2), 1, 1) = "*" Then
           ts(2) = Mid$(ts(1), 2, Len(ts(0)) - 1)
        End If
        If ts(0) = "" Or ts(0) = "+" Or ts(0) = "#" Then
           ts(0) = "1"
        ElseIf ts(0) = "-" Or ts(0) = "@" Then
           ts(0) = "-1"
        End If
        If ts(2) = "" Or ts(2) = "+" Or ts(2) = "#" Then
           ts(2) = "1"
        ElseIf ts(2) = "-" Or ts(2) = "@" Then
           ts(2) = "-1"
        End If
         set_string = time_string(set_string(ts(0)), set_string(ts(1)), True, False)
         set_string = time_string(set_string, set_string(ts(2)), True, False)
    End If
    End If
   End If
  End If
End If
End Function
Private Function simple_string0(s As String, factor As String) As String
Dim S_p As para_item_type
Dim f(1) As String
Dim ts As String
Dim i%
On Error GoTo simple_string0_error
If InStr(1, s, "F", 0) > 0 Then
 simple_string0 = "F"
  Exit Function
End If
If read_string(s, S_p) = False Then
 simple_string0 = s
  factor = "1"
Else
If S_p.last_it = 1 Or S_p.last_it = 0 Then
 ts = Mid$(s, 1, 1)
 If ts = "-" Or ts = "@" Then
  factor = "@1"
  s = Mid$(s, 2, Len(s) - 1)
  simple_string0 = "-" + s
 ElseIf ts = "+" Or ts = "#" Then
  s = Mid$(s, 2, Len(s) - 1)
  factor = "1"
  simple_string0 = s
 Else
  simple_string0 = s
  factor = "1"
 End If
   Exit Function
End If
Call simple_multi_para_for_sp(S_p, f(0))
Call simple_multi_item_for_sp(S_p, f(1))
If f(0) = "1" And f(1) = "1" Then
 simple_string0 = s
  factor = "1"
   Exit Function
End If
 If f(0) = "1" Then
 factor = f(1)
 ElseIf f(1) = "1" Then
 factor = f(0)
 Else
 factor = f(0) + "*" + f(1)
 End If
simple_string0 = ""
For i% = 0 To S_p.last_it - 1
  ts = time_string(S_p.pA(i%), S_p.it(i%), True, False)
   If simple_string0 = "" And ts <> "" Then
    simple_string0 = ts
   ElseIf simple_string0 <> "" And ts <> "" Then
    If Mid$(ts$, 1, 1) = "-" Then
    simple_string0 = simple_string0 + ts
    Else
    simple_string0 = simple_string0 + "+" + ts
    End If
   End If
Next i%
 If InStr(1, simple_string0, "F", 0) = 0 Then
 s = simple_string0
 Else
 simple_string0 = s
 factor = "1"
 End If
End If
Exit Function
simple_string0_error:
 simple_string0 = s
 factor = "1"
'If factor = "-1" Then
 'simple_string0 = "-" + simple_string0
'ElseIf factor <> "1" Then
 'simple_string0 = factor + simple_string0
'End If
End Function

Public Sub unsimple_string(s As String)
Dim i%, j%, k%, l%, p%, m%, n%
Dim S_p As para_item_type
Dim ts(2) As String
If InStr(1, s, "/", 0) > 0 Then ' 分式
 Exit Sub
End If
Call remove_brace(s)
If s = "" Then
 Exit Sub
End If
If Mid$(s, Len(s), 1) = ")" Then
 p% = 1
  For i% = Len(s) - 1 To 1 Step -1
   If Mid$(s, i%, 1) = ")" Then
    p% = p% + 1
   ElseIf Mid$(s, i%, 1) = "(" Then
    p% = p% - 1
     If p% = 0 Then
      n% = i%
      GoTo unsimple_string_mark1
      End If
   End If
  Next i%
unsimple_string_mark1:
If n% > 1 Then
 If Mid$(s, n% - 1, 1) = "'" Then
  Exit Sub
 End If
End If
For i% = n% - 1 To 1 Step -1
If Mid$(s, i%, 1) = "+" Or Mid$(s, i%, 1) = "-" Then
 m% = i%
  GoTo unsimple_string_mark2
End If
Next i%
If Mid$(s, Len(s), 1) <= "A" Then
s = unsimple_para(s)
End If
unsimple_string_mark2:
If m% = 1 Then
   ts(0) = Mid$(s, 1, n% - 1)
    ts(1) = Mid$(s, n%, Len(s) - n% + 1)
     If ts(0) = "-" Then
      ts(0) = "-1"
     ElseIf ts(0) = "+" Then
      ts(0) = "1"
     End If
     If Mid$(ts(1), 1, 1) = "*" Then
      ts(1) = Mid$(ts(1), 2, Len(ts(1)) - 1)
      End If
Call remove_brace(ts(1))
If read_string(ts(1), S_p) Then
 s = ""
 For i% = 0 To S_p.last_it - 1
  S_p.pA(i%) = time_para(S_p.pA(i%), ts(0), True, False)
   ts(2) = time_string(S_p.pA(i%), S_p.it(i%), False, False)
   If s = "" And ts(2) <> "" Then
    s = ts(2)
   ElseIf s <> "" And ts(2) <> "" Then
    s = s + "+" + ts(2)
   End If
 Next i%
'Else

 End If
 End If
End If
End Sub


Private Function combine_string_from_para_item(ByVal p1$, _
    ByVal p2$, ByVal p3$, ByVal p4$, ByVal i1$, ByVal i2$, ByVal i3$, _
         ByVal I4$, n%) As String
         '组合为多项式
Dim i%
Dim ts(3) As String
If InStr(1, p1$, "F", 0) > 0 Or InStr(1, p2$, "F", 0) > 0 Or _
        InStr(1, p3$, "F", 0) > 0 Or InStr(1, p4$, "F", 0) > 0 Then
 combine_string_from_para_item = "F"
  Exit Function
End If
ts(0) = time_string(p1$, i1$, True, False)
ts(1) = time_string(p2$, i2$, True, False)
ts(2) = time_string(p3$, i3$, True, False)
ts(3) = time_string(p4$, I4$, True, False)
combine_string_from_para_item = ts(0)
For i% = 1 To n% - 1
 If ts(i%) <> "" And ts(i%) <> "0" Then
  If Mid$(ts(i%), 1, 1) = "-" Then
  combine_string_from_para_item = combine_string_from_para_item + _
   ts(i%)
  Else
  combine_string_from_para_item = combine_string_from_para_item + _
   "+" + ts(i%)
  End If
 End If
Next i%
End Function

Private Function simple_multi_long(n1&, n2&, n3&, n4&, n5&, n6&, n7&, n8&, n9&, n%, n0&) As Boolean
Dim i%, j%, k%
Dim tn(8) As Long
Dim tn1(8) As Long
Dim t_n&
On Error GoTo simple_multi_long_error
simple_multi_long = True
If n1& <> 0 Then
n0& = n1&
Else
n0& = 1
End If
'k% = n1%
If n1& > 0 Then
tn(0) = n1&
tn(1) = n2&
tn(2) = n3&
tn(3) = n4&
tn(4) = n5&
tn(5) = n6&
tn(6) = n7&
tn(7) = n8&
tn(8) = n9&
ElseIf n1& < 0 Then
tn(0) = -n1&
tn(1) = -n2&
tn(2) = -n3&
tn(3) = -n4&
tn(4) = -n5&
tn(5) = -n6&
tn(6) = -n7&
tn(7) = -n8&
tn(8) = -n9&
Else
 If n1& <> 0 Then
 n0& = n0& / n1&
 End If
 Exit Function
End If
t_n& = tn(0)
 For j% = 1 To n%
  t_n& = l_gcd(t_n&, tn(j%))
Next j%

n1& = tn(0) / t_n&
n2& = tn(1) / t_n&
n3& = tn(2) / t_n&
n4& = tn(3) / t_n&
n5& = tn(4) / t_n&
n6& = tn(5) / t_n&
n7& = tn(6) / t_n&
n8& = tn(7) / t_n&
n9& = tn(8) / t_n&
If n1& <> 0 Then
n0& = n0& / n1&
End If
Exit Function
simple_multi_long_error:
simple_multi_long = False
End Function
Public Function do_factor1(ByVal s As String, mp$, f1$, f2$, f3$, n%) As Boolean
'n% 因子数
Dim p_i As para_item_type
Dim t_p(3) As String
Dim t_i(3) As String
Dim ts As String
Dim ts_ As String
Dim i%, j%, k%, l%
 mp$ = s
  f1$ = "1"
   f2$ = "1"
    f3 = "1"
     n% = 1
If InStr(1, s, "[", 0) > 0 Then '含重根号
 mp$ = s
  f1$ = "1"
   f2$ = "1"
    f3 = "1"
     n% = 1
 do_factor1 = True
  Exit Function
End If
'******************************************************
n% = 4
If InStr(1, s, "/", 0) > 0 Then '分式
do_factor1_error:
 do_factor1 = False
  Exit Function
End If
'*********************************************************
i% = InStr(1, s, "(", 0) '确定第一个符号l%+-;
j% = InStr(1, s, "'", 0)
k% = InStr(2, s, "+", 0)
l% = InStr(2, s, "-", 0)
If k% > 0 And l% > 0 Then
 If k% < l% Then
    l% = k% '+-号
 End If
ElseIf l% = 0 And k% > 0 Then
    l% = k%
End If
k% = InStr(2, s, "#", 0)
If k% > 0 And l% > 0 Then
 If k% < l% Then
    l% = k%
 End If
ElseIf l% = 0 And k% > 0 Then
    l% = k%
End If
k% = InStr(2, s, "@", 0)
If k% > 0 And l% > 0 Then
 If k% < l% Then
    l% = k%
 End If
ElseIf l% = 0 And k% > 0 Then
    l% = k%
End If
'******************************
If i% >= 2 And (i% < j% Or j% = 0) And i% < l% Then '(第一
 If Mid$(s, Len(s), 1) = ")" Then
  mp$ = Mid$(s, 1, i% - 1)
  If mp$ = "-" Then
   mp$ = "-1"
  ElseIf mp$ = "+" Then
   mp$ = "1"
  End If
  do_factor1 = do_factor1(Mid$(s, i%, Len(s) - i% + 1), f1$, f2$, f3$, "", n%)
   If mp$ <> "1" Then
   n% = n% + 1
   End If
  do_factor1 = True
  Exit Function
 End If
'ElseIf j% > 2 And (j% < i% Or i% = 0) Then
'   mp$ = Mid$(s, 1, j% - 1)
 ' If mp$ = "-" Then
  ' mp$ = "-1"
  'ElseIf mp$ = "+" Then
 '  mp$ = "1"
 ' End If
 ' f1$ = Mid$(s, i%, Len(s) - i% + 1)
 ' f2$ = "1"
 ' f3$ = "1"
 ' do_factor1 = True
 ' Exit Function

End If
'******************
mp$ = s
f1$ = "1"
f2$ = "1"
f3$ = "1"
On Error GoTo do_factor1_error
'If p_i.last_it < 4 Then
If InStr(1, s, "+", 0) < 2 And InStr(1, s, "-", 0) < 2 _
    And InStr(1, s, "@", 0) < 2 And InStr(1, s, "#", 0) < 2 Then
 For i% = 1 To Len(s)
 ts = Mid$(s, i%, 1)
  If Asc(ts) < 0 And ts <> "'" And ts <> "\" Then
   t_p(0) = Mid$(s, 1, i% - 1)
   t_p(2) = Mid$(s, i% + 1, Len(s) - i%)
     If Len(t_p(0)) > 1 Then
      If Mid$(t_p(0), Len(t_p(0)), 1) = "*" Then
       t_p(0) = Mid$(t_p(0), 1, Len(t_p(0)) - 1)
      End If
     End If
     If Len(t_p(1)) > 1 Then
      If Mid$(t_p(1), 1, 1) = "*" Then
       t_p(1) = Mid$(t_p(1), 2, Len(t_p(0)) - 1)
      End If
     End If
      If t_p(0) = "" And t_p(1) = "" Then
       mp$ = ts
      Else
       mp$ = ts
       If t_p(0) = "" Then
        Call do_factor1(t_p(1), f1$, f2$, f3$, "", 0)
       ElseIf t_p(1) = "" Then
        f1$ = t_p(0)
       Else
        Call do_factor1(time_string(t_p(0), t_p(1), True, False), f1$, f2$, f3$, "", 0)
       End If
      End If
            If mp$ = "1" Then
             n% = n% - 1
              mp$ = f1$
               f1$ = f2$
                f2$ = f3$
                 f3$ = "1"
            End If
            If f1$ = "1" Then
             n% = n% - 1
              f1$ = f2$
               f2$ = f3$
                f3$ = "1"
            End If
            If f2$ = "1" Then
             n% = n% - 1
              f2$ = f3$
               f3$ = 1
            End If
            If f3$ = "1" Then
             n% = n% - 1
            End If
      Exit Function
    ElseIf ts = "'" Then
      If Mid$(s, i% + 1, 1) = "(" Then
       k% = 1
        For j% = i% + 2 To Len(s)
         ts_ = Mid$(s, j%, 1)
         If ts_ = "(" Then
          k% = k% + 1
         ElseIf ts_ = ")" Then
          k% = k% - 1
           If k% = 0 Then
            t_p(0) = Mid$(s, 1, i% - 1)
            t_p(1) = Mid$(s, i%, j% - i% + 1)
            t_p(2) = Mid$(s, j% + 1, Len(s) - j%)
            mp$ = t_p(0)
            f1$ = t_p(1)
            Call do_factor1(t_p(2), f2$, f3$, "", "", 0)
             do_factor1 = True
            If mp$ = "1" Then
             n% = n% - 1
              mp$ = f1$
               f1$ = f2$
                f2$ = f3$
                 f3$ = "1"
            End If
            If f1$ = "1" Then
             n% = n% - 1
              f1$ = f2$
               f2$ = f3$
                f3$ = "1"
            End If
            If f2$ = "1" Then
             n% = n% - 1
              f2$ = f3$
               f3$ = 1
            End If
            If f3$ = "1" Then
             n% = n% - 1
            End If
             Exit Function
           End If
         End If
        Next j%
      Else
        For j% = i% + 1 To Len(s)
         ts_ = Mid$(s, j%, 1)
         If ts_ >= "A" Then
          t_p(0) = Mid$(s, 1, j% - 1)
          t_p(2) = Mid$(s, j%, Len(s) - j% + 1)
            mp$ = t_p(0)
            Call do_factor1(t_p(2), f1$, f2$, f3$, "", 0)
            do_factor1 = True
            If mp$ = "1" Then
             n% = n% - 1
              mp$ = f1$
               f1$ = f2$
                f2$ = f3$
                 f3$ = "1"
            End If
            If f1$ = "1" Then
             n% = n% - 1
              f1$ = f2$
               f2$ = f3$
                f3$ = "1"
            End If
            If f2$ = "1" Then
             n% = n% - 1
              f2$ = f3$
               f3$ = 1
            End If
            If f3$ = "1" Then
             n% = n% - 1
            End If
            Exit Function
         ElseIf j% = Len(s) Then
           mp$ = s
          f1$ = "1"
          f2$ = "1"
          f3$ = "1"
             do_factor1 = True
             n% = 1
             Exit Function
          End If
         Next j%
     End If
    End If
    Next i%
 mp$ = s
Else
do_factor1 = True
 If read_string(s, p_i) = False Then
  mp$ = s
   f1$ = "1"
    f2$ = "1"
     f3$ = "1"
      Exit Function
 End If
 If p_i.last_it > 3 Then 'Or p_i.last_it < 2 Then
  do_factor1 = False
  mp$ = s
   f1$ = "1"
    f2$ = "1"
     f3$ = "1"
   Exit Function
 ElseIf p_i.last_it = 1 Then
  do_factor1 = False
 For i% = 0 To Len(s)
 ts = Mid$(s, 1, i%)
 If ts <> "*" Then
 mp$ = ""
 f3$ = ""
 If ts < "a" Then
 mp$ = mp$ + ts
 Else
 f3$ = f3$ + ts
 End If
End If
Next i%
 If f3$ = "" Then
  f3$ = "1"
 End If
  If mp$ = "" Then
  mp$ = "1"
  ElseIf mp$ = "-" Then
  mp$ = "-1"
  End If
 Else
  For i% = 0 To p_i.last_it - 1
    t_p(i%) = p_i.pA(i%)
     t_i(i%) = p_i.it(i%)
  Next i%
 do_factor1 = do_factor(t_p(0), t_p(1), t_p(2), t_p(3), _
       t_i(0), t_i(1), t_i(2), t_i(3), p_i.last_it, _
         mp$, f1$, f2$, f3$)
  If do_factor1 = False Then
     mp$ = "1"
      f1$ = s
       f2$ = "1"
        f3$ = "1"
  End If
End If
End If
If Mid$(mp$, 1, 1) = "@" Or Mid$(mp$, 1, 1) = "-" Then
If f3$ <> "" And f3$ <> "1" Then
 mp$ = time_para(mp$, "@1", True, False)
 f3$ = time_string(f3$, "-1", True, False)
ElseIf f2$ <> "" And f2$ <> "1" Then
 mp$ = time_para(mp$, "@1", True, False)
 f2$ = time_string(f2$, "-1", True, False)
ElseIf f1$ <> "" And f1$ <> "1" Then
 mp$ = time_para(mp$, "@1", True, False)
 f1$ = time_string(f1$, "-1", True, False)
End If
End If
            If f1$ = "1" Then
             n% = n% - 1
              f1$ = f2$
               f2$ = f3$
                f3$ = "1"
            End If
            If f2$ = "1" Then
             n% = n% - 1
              f2$ = f3$
               f3$ = 1
            End If
            If f3$ = "1" Then
             n% = n% - 1
            End If
End Function

Private Function simple_multi_para_for_sp(S_p As para_item_type, p0$) As Boolean
Dim i%, j%
Dim fac() As String
Dim tn() As Long
Dim tn1() As Long
Dim t_n(1) As Long
Dim t_n1(1) As Long
Dim temp_pi As para_item_type
On Error GoTo simple_multi_para_for_sp_error
simple_multi_para_for_sp = True
p0$ = "1"
If S_p.last_it < 2 Then
 Exit Function
End If
If S_p.last_it > 0 Then
  p0$ = S_p.pA(0)
Else
  Exit Function
End If
For i% = 0 To S_p.last_it - 1
 If S_p.pA(i%) = "1" Then
  p0$ = "1"
  Exit Function
 End If
Next i%
t_n(0) = 0
t_n1(0) = 0
For i% = 0 To S_p.last_it - 1
Call read_para(S_p.pA(i%), temp_pi)
For j% = 0 To temp_pi.last_it - 1
ReDim Preserve tn(j%) As Long
ReDim Preserve tn1(j%) As Long
 tn(j%) = val_(temp_pi.pA(j%))
 tn1(j%) = val_(temp_pi.it(j%))
Next j%
Call simple_multi_long1(tn(), temp_pi.last_it, t_n(1))
Call simple_multi_long1(tn1(), temp_pi.last_it, t_n1(1))
If t_n(0) = 0 Then
 t_n(0) = t_n(1)
Else
 Call simple_two_long(t_n(0), t_n(1), t_n(0))
End If
If t_n1(0) = 0 Then
 t_n1(0) = t_n1(1)
Else
 Call simple_two_long(t_n1(0), t_n1(1), t_n1(0))
End If
Next i%
For i% = 0 To S_p.last_it - 1
 Call read_para(S_p.pA(i%), temp_pi)
  For j% = 0 To temp_pi.last_it - 1
   temp_pi.pA(j%) = str_(val_(temp_pi.pA(j%)) / val(t_n(0)))
   temp_pi.it(j%) = str_(val_(temp_pi.it(j%)) / val(t_n1(0)))
  Next j%
 S_p.pA(i%) = para_from_para_item(temp_pi)
Next i%
p0$ = combine_para_for_item(str_(t_n(0)), str_(t_n1(0)))
Exit Function
simple_multi_para_for_sp_error:
simple_multi_para_for_sp = False
End Function

Private Function simple_multi_long1(tn() As Long, n As Integer, n0 As Long) As Boolean
Dim i%, j%, k%
Dim tn1(8) As Long
Dim t_n&
On Error GoTo simple_multi_long1_error
simple_multi_long1 = True
If n% > 1 Then
If tn(0) <> 0 Then
n0& = tn(0)
Else
n0& = 1
End If
ElseIf n% = 1 Then
n0 = tn(0)
tn(0) = 1
Exit Function
Else
Exit Function
End If
'k% = n1%
If tn(0) < 0 Then
For i% = 0 To n% - 1
tn(i%) = -tn(i%)
Next i%
End If

 'If tn(0) <> 0 Then
 'n0& = n0& / tn(0)
 'End If
t_n& = tn(0)
 For j% = 1 To n% - 1
  t_n& = l_gcd(t_n&, tn(j%))
Next j%
For i% = 0 To n% - 1
tn(i%) = tn(i%) / t_n&
Next i%
If tn(0) <> 0 Then
n0& = n0& / tn(0)
End If
Exit Function
simple_multi_long1_error:
simple_multi_long1 = False
End Function

Private Function simple_multi_item_for_sp(S_p As para_item_type, I0$) As Boolean
Dim i%, j%
Dim ch As String * 1
Dim tn() As Integer
Dim ts$
On Error GoTo simple_multi_item_for_sp_error
simple_multi_item_for_sp = True
'If s_p.last_it > 0 Then
I0$ = "1"
If S_p.last_it < 2 Then
 Exit Function
End If
If S_p.last_it > 0 Then
ts$ = S_p.it(0)
End If
For i% = 0 To S_p.last_it - 1
 If InStr(1, S_p.it(i%), "'", 0) > 0 Or InStr(1, S_p.it(i%), "[", 0) > 0 Then
  Exit Function
 End If
 ReDim Preserve tn(i%) As Integer
 If S_p.it(i%) = "1" Then
  Exit Function
 End If
Next i%
i% = 1
Do
simple_multi_item_mark1:
If Len(S_p.it(0)) >= i% Then
ch = Mid$(S_p.it(0), i%, 1)
Else
GoTo simple_multi_item_mark2
End If
For j% = 1 To S_p.last_it - 1
 tn(j%) = InStr(1, S_p.it(j%), ch, 0)
  If tn(j%) = 0 Then
   i% = i% + 1
    GoTo simple_multi_item_mark1
  End If
Next j%
S_p.it(0) = Mid$(S_p.it(0), 1, i% - 1) + Mid$(S_p.it(0), i% + 1, Len(S_p.it(0)) - i%)
For j% = 1 To S_p.last_it - 1
 S_p.it(j%) = Mid$(S_p.it(j%), 1, tn(j%) - 1) + Mid$(S_p.it(j%), tn(j%) + 1, Len(S_p.it(j%)) - tn(j%))
  If S_p.it(j%) = "" Then
     S_p.it(j%) = "1"
  End If
Next j%
Loop
simple_multi_item_mark2:
I0$ = divide_item(ts$, S_p.it(0), "", "")
Exit Function
simple_multi_item_for_sp_error:
simple_multi_item_for_sp = False
End Function

Public Sub read_brace(s As String, st%, n1%, n2%)
' 读()
Dim p%, i%
n1% = 0
 n2% = 0
  p% = 0
For i% = st% To Len(s)
 If Mid$(s, i%, 1) = "(" Then
  n1% = i%
   p% = p% + 1
  ElseIf Mid$(s, i%, 1) = ")" Then
   p% = p% - 1
    If p% = 0 Then
     n2% = i%
      Exit Sub
    End If
   End If
Next i%
 n1% = 0
  n2% = 0
End Sub

Public Sub read_item1(s1$, S2$, s3$)
Dim i%
i% = InStr(1, s1, "'", 0)
If i% = 0 Then
 S2$ = s1$
 s3$ = "1"
ElseIf i% = 1 Then
S2$ = "1"
s3$ = Mid$(s1, 2, Len(s1) - 1)
ElseIf i% > 1 Then
 S2$ = Mid$(s1, 1, i% - 1)
 If S2 = "" Then
  S2 = "1"
 End If
 s3$ = Mid$(s1, i% + 1, Len(s1) - i%)
End If
End Sub

Private Sub read_item0(s1$, S2$, s3$)
Dim i%, j%
j% = InStr(1, s1$, "'")
i% = InStr(1, s1, "*")
If i% > j% Then
 i% = 0
End If
If i% = Len(s1$) Then
 S2$ = Mid$(s1$, 1, Len(s1$) - 1)
  s3$ = "1"
ElseIf i% > 0 Then
 S2 = Mid$(s1$, 1, i% - 1)
  s3 = Mid$(s1$, i% + 1, Len(s1$) - i%)
  If S2 = "-" Then
   S2 = "-1"
  End If
Else
 If Mid$(s1$, 1, 1) = "-" Then
  S2$ = "-1"
   s3$ = Mid$(s1$, 2, Len(s1$) - 1)
 Else
 S2$ = "1"
  s3$ = s1$
 End If
End If
End Sub

Private Function is_in_brace(s As String, m%) As Boolean
Dim i%, k%
k% = 0
For i% = m% - 1 To 1 Step -1
 If Mid$(s, i%, 1) = "(" Then
  k% = k% - 1
   If k% = -1 Then
    is_in_brace = True
     Exit Function
    End If
 ElseIf Mid$(s, i%, 1) = ")" Then
  k% = k + 1
 End If
Next i%
  is_in_brace = False
End Function

Private Sub read_item(s1$, S2$, s3$)
Dim m%, i%, n%, A%
Dim ts$
S2$ = ""
 s3$ = ""
m% = InStr(1, s1$, "*", 0)
If m% = 0 Then
    If get_brace_pair(s1$, 1, m%, n%) Then
     ts$ = Mid$(s1$, m%, n% - m% + 1)
      If is_great_than_A(ts$, A%) Then
       If m% > 1 Then
        If Mid$(s1$, m% - 1, 1) = "'" Then
         s3$ = "'" + ts
          If m% > 2 Then
           S2$ = Mid$(s1, 1, m - 2)
            If S2$ = "" Then
             S2$ = "1"
            ElseIf S2$ = "-" Then
             S2$ = "-1"
            End If
          Else 'm%=2
           S2$ = "1"
          End If
         Else '<>"'"
          S2$ = Mid$(s1, 1, m - 1)
           If S2$ = "" Then
             S2$ = "1"
           ElseIf S2$ = "-" Then
             S2$ = "-1"
           End If
         End If
       Else 'm%=1
        s3$ = ts$
         S2$ = "1"
       End If
     Else 'ts<"A"false
      S2$ = Mid$(s1, 1, n%)
       If n% < Len(s1$) Then
        s3$ = Mid$(s1, n%, Len(s1$) - n%)
       Else
        s3$ = "1"
       End If
     End If
  Else 'no_brace
    If is_great_than_A(s1$, A%) Then
      If Mid$(s1$, A%, 1) > "A" Then
       If A% = 1 Then
        S2$ = "1"
         s3$ = s1$
          Exit Sub
       ElseIf A% >= 2 Then
         If Mid$(s1, A% - 1, 1) = "'" Then
           s3$ = Mid$(s1$, A% - 1, Len(s1) - A% + 2)
            If A% > 2 Then
            S2$ = Mid$(s1, 1, A% - 2)
             If S2$ = "-" Then
              S2$ = "-"
              End If
            Else
            S2$ = "1"
            End If
             Exit Sub
         Else '<>"'"
            s3$ = Mid$(s1$, A%, Len(s1) - A% + 1)
             S2$ = Mid$(s1$, 1, A% - 1)
              If S2$ = "-" Then
               S2$ = "-1"
              End If
              Exit Sub
         End If
       End If
      End If
    Else
     S2$ = s1$
      s3$ = "1"
    End If
  End If
 Else
  S2$ = Mid$(s1, 1, m% - 1)
   s3$ = Mid$(s1, m% + 1, Len(s1) - m%)
    If s3$ = "" Then
     s3$ = "1"
    End If
    If S2$ = "" Then
     S2$ = "1"
    ElseIf S2$ = "-" Then
     S2$ = "-1"
    End If
 End If

End Sub


Public Function gcd_for_string(ByVal s1$, ByVal S2$, gcd$, f1$, f2$, is_simple As Boolean) As Boolean
Dim f(7) As String
Dim i%, j%
gcd$ = "1"
f1$ = "1"
f2$ = "1"
On Error GoTo gcd_for_string_error
f1$ = s1$
f2$ = S2$
Call simple_two_string_(f1$, f2$, gcd$)
If gcd <> "1" Then
 gcd_for_string = True
Else
 gcd_for_string = False
 gcd$ = "1"
 f1$ = s1$
 f2$ = S2$
End If
Exit Function
gcd_for_string_error:
gcd_for_string = False
End Function

Public Function lcd_for_string(ByVal s1$, ByVal S2$, lcd$, f1$, f2$, is_simple As Boolean) As Boolean '表达式开方
Dim ts(7) As String
Dim i%, j%
Dim tn(2) As Long
Dim ty(1) As Byte
'最小共倍式
For i% = 1 To Len(s1$)
 If Mid$(s1$, i%, 1) > "9" Or Mid$(s1$, i%, 1) < "0" Then
  ty(0) = 1
   GoTo lcd_for_string_next1
 End If
Next i%
lcd_for_string_next1:
For i% = 1 To Len(S2$)
 If Mid$(S2$, i%, 1) > "9" Or Mid$(S2$, i%, 1) < "0" Then
  ty(1) = 1
   GoTo lcd_for_string_next2
 End If
Next i%
lcd_for_string_next2:
If ty(0) = 1 And ty(1) = 1 Then
If gcd_for_string(s1, S2, ts(0), ts(1), ts(2), True) Then
lcd = time_string(ts(0), time_string(ts(1), ts(2), False, False), True, False)
f1$ = ts(1)
f2$ = ts(2)
lcd_for_string = True
Else
lcd = time_string(s1$, S2$, True, False)
f1$ = s1$
f2$ = S2$
If InStr(1, lcd, "F", 0) = 0 Then
lcd_for_string = True
End If
End If
ElseIf ty(0) = 0 And ty(1) = 0 Then
  tn(0) = val(s1$)
  tn(1) = val(S2$)
tn(2) = l_gcd(tn(0), tn(1))
tn(0) = tn(0) / tn(2)
tn(1) = tn(1) / tn(2)
f1$ = Trim(str(tn(0)))
f2$ = Trim(str(tn(1)))
lcd$ = Trim(str(tn(0) * tn(2) * tn(1)))
lcd_for_string = True
ElseIf ty(0) = 0 Then
 S2$ = simple_string(S2$)
  j% = InStr(1, S2$, "(", 0)
  If j% > 1 Then
     ts(0) = Mid(S2$, i%, j% - 1)
      ts(1) = Mid(S2$, j%, Len(S2$) - j% + 1)
  Else
     ts(0) = "1"
      ts(1) = S2$
  End If
   ty(0) = 0
  For i% = 1 To Len(ts(0))
   If Mid$(ts(0), i%, 1) >= "9" Or Mid$(ts(0), i%, 1) <= "0" Then
    ty(0) = 1
     GoTo lcd_for_string_next3
   End If
  Next i%
lcd_for_string_next3:
If ty(0) Then
tn(0) = val(s1$)
tn(1) = val(ts(0))
tn(2) = l_gcd(tn(0), tn(1))
tn(0) = tn(0) / tn(2)
tn(1) = tn(1) / tn(2)
f1$ = Trim(str(tn(0)))
f2$ = Trim(str(tn(1)))
f2$ = time_string(ts(1), f2$, True, False)
lcd$ = time_string(time_string(f1$, f2$, False, False), Trim(str(tn(2))), True, False)
If InStr(1, lcd$, "F", 0) = 0 Then
lcd_for_string = True
End If
Else
End If
ElseIf ty(1) = 0 Then
lcd_for_string = lcd_for_string(S2$, s1$, lcd$, f2$, f1$, is_simple)
End If
End Function
Public Function is_pure_number(ByVal s As String) As Boolean '判断是否为纯数
Dim i%, la%
Dim ch$
For i% = 1 To Len(s)
ch$ = Mid$(s, i%, 1)
If ch$ >= "A" And ch$ <= "z" Then
 is_pure_number = False
  Exit Function
ElseIf ch$ = "[" Then
   is_pure_number = _
     is_pure_number(number_string(read_sqr_no_from_string(s, i%, i%, "")))
   If is_pure_number = False Then
    Exit Function
   End If
End If
Next i%
is_pure_number = True
End Function

Private Sub rational_item(s1$, S2$, s3$)
Dim n%, m%
S2$ = "1"
 s3$ = "1"
n% = InStr(1, s1$, "'", 0)
If n% = 0 Then
 S2$ = "1"
  s3$ = s1$
 Exit Sub
Else
 If Mid$(s1$, n% + 1, 1) = "(" Then
     m% = right_brace(s1$, n% + 2)
 S2$ = Mid$(s1$, n%, m% - n% + 1)
 s3$ = Mid$(s1$, 1, n% - 1)
  If Len(s1$) > n% Then
      s3$ = s3$ + Mid$(s1$, n% + 1, Len(s1$) - n%)
  End If
 Else
 S2$ = "'" + Mid$(s1$, n% + 1, Len(s1$) - n%)
 s3$ = Mid$(s1$, 1, n% - 1) + _
         Mid$(s1$, n% + 1, Len(s1$) - n%)
 End If
 If s3$ = "" Or s3$ = "+" Or s3$ = "#" Then
  s3$ = "1"
 ElseIf s3$ = "@" Or s3$ = "-" Then
  s3$ = "-1"
 End If
 If S2$ = "" Or S2$ = "+" Or S2$ = "#" Then
  S2$ = "1"
 ElseIf S2$ = "@" Or S2$ = "-" Then
  S2$ = "-1"
 End If
End If
End Sub

Public Function value_string(ByVal s1$) As String '将表达式化为数
Dim ty As Byte
Dim S2$, s3$, S4$, S5$, S6$, S7$
ty = string_type(s1$, "", S2$, s3$, S4$)
If ty = 0 Then
 If S4$ = "" Then
  If s3$ <> "1" Then
  value_string = value_para(S2$) + s3$
  Else
  value_string = value_para(S2$)
  End If
 Else
     S2$ = value_para(S2$)
     S4$ = value_string(S4$)
  Call string_type(S4$, "", S5$, S6$, S7$)
   If S6$ = s3$ Then
    S2$ = str_(val_(S2$) + val_(S5$))
     S4$ = S7$
   End If
  If S4$ <> "" Then
  If s3 <> "1" Then
   value_string = S2$ + s3$ + "+" + S4$
  Else
   value_string = S2$ + "+" + S4$
  End If
  Else
  If s3$ <> "1" Then
  value_string = S2$ + s3$
  Else
  value_string = S2$
  End If
  End If
 End If
ElseIf ty = 1 Then
ElseIf ty = 2 Then
ElseIf ty = 3 Then
 value_string = str_(val(value_string(S2$)) / val(value_string(s3$)))
End If
End Function

Public Function simple_string1(ByVal s$) As String
'整理多项式
Dim ty As Byte
Dim ts(3) As String
On Error GoTo simple_string1_error
If InStr(1, s$, "F", 0) > 0 Then
 GoTo simple_string1_error
End If
ty = string_type(s$, ts(0), ts(1), ts(2), ts(3))
If ty = 3 Then
 simple_string1 = add_brace(simple_string1(ts(1)), _
       "") + _
      "/" + add_brace(simple_string1(ts(2)), _
         "root")
ElseIf ty = 0 Then
 If ts(3) = "" Then
  simple_string1 = s$
 Else
  simple_string1 = add_string(ts(0), simple_string1(ts(3)), True, False)
 End If
Else
 simple_string1 = s$
End If
Exit Function
simple_string1_error:
simple_string1 = "F"
End Function

Public Function get_brace_pair(ByVal s$, ByVal start%, n1%, n2%) As Boolean '获取配对括号
Dim i%, j%
Dim ts As String
n1% = 0
n2% = 0
n1% = InStr(start%, s$, "(")
If n1% > 0 Then
j% = 1
For i% = n1% + 1 To Len(s$)
 ts = Mid$(s$, i%, 1)
  If ts = ")" Then
   j% = j% - 1
  ElseIf ts = "(" Then
   j% = j% + 1
  End If
  If j% = 0 Then
   get_brace_pair = True
    n2% = i%
     Exit Function
  End If
Next i%
Else
get_brace_pair = False
End If
End Function

Public Function is_great_than_A(ByVal s$, n%) As Boolean
Dim i%
n% = 0
For i% = 1 To Len(s$)
 If Mid$(s$, i%, 1) > "A" Then
  is_great_than_A = True
   n% = i%
   Exit Function
 End If
Next i%
End Function

Private Sub read_item_from_string(ByVal s$, s1$, S2$)
Dim m%, n%, f_brace%, e_brace%
Dim ty As Byte
Dim ts(6) As String
Dim it(1) As String
Call unsimple_string(s$)
m% = InStr(2, s$, "+", 0)
 n% = InStr(2, s$, "-", 0)
  If (n% > 0 And n% < m%) Or m% = 0 Then
   m% = n%
    n% = InStr(2, s$, "#", 0)
     If (n% > 0 And n% < m%) Or m% = 0 Then
      m% = n%
       n% = InStr(2, s$, "@", 0)
     If (n% > 0 And n% < m%) Or m% = 0 Then
      m% = n%
     End If
  End If
 End If
  f_brace% = 0
  e_brace% = 0
read_item_from_string_back:
 If m% > 0 Then
  Call get_brace_pair(s$, e_brace% + 1, f_brace%, e_brace%)
   If m% > e_brace% Or m% < f_brace% Then ' ()外
   s1$ = Mid$(s$, 1, m% - 1)
    S2$ = Mid$(s$, m%, Len(s$) - m% + 1)
     If Mid$(S2$, 1, 1) = "+" Then
      S2$ = Mid$(S2$, 2, Len(S2$) - 1)
     End If
   Else
    m% = InStr(e_brace% + 1, s$, "+", 0)
    n% = InStr(e_brace% + 1, s$, "-", 0)
    If (n% > 0 And n% < m%) Or m% = 0 Then
     m% = n%
    End If
    GoTo read_item_from_string_back
   End If
 Else
   s1 = s$
    S2$ = ""
 End If
End Sub

Public Function solve_equation_group(ByVal E_no%, Equ As Equation_group_type, cal_float As Boolean) As Boolean   '解一次方程
Dim r1$, r2$
If Equ.data(0).equation(0) > 0 And Equ.data(0).equation(1) = 0 Then
   solve_equation_group = solve_equation(equation(Equ.data(0).equation(0)).data(0), r1$, r2$, cal_float)
    
Else
End If
End Function

Public Function solve_equation(Equ As Equation_data0_type, r1$, r2$, cal_float As Boolean) As Boolean  '解一次方程
Dim temp_record As total_record_type
temp_record.record_data = Equ.record
If Equ.para_xx <> "0" Then
 solve_equation = solut_2order_equation(Equ.para_xx, Equ.para_x, Equ.para_c, _
                   r1$, r2$, cal_float)
ElseIf Equ.para_yy <> "0" Then
 solve_equation = solut_2order_equation(Equ.para_yy, Equ.para_y, Equ.para_c, _
                   r1$, r2$, cal_float)
ElseIf Equ.para_x <> "0" Then
 r1$ = divide_string(time_string("-1", Equ.para_c, True, cal_float), Equ.para_x, _
        True, cal_float)
        r2$ = ""
        solve_equation = True
ElseIf Equ.para_y <> "0" Then
 r1$ = divide_string(time_string("-1", Equ.para_c, True, cal_float), Equ.para_y, _
        True, cal_float)
        r2$ = ""
        solve_equation = True
Else
End If
End Function

Public Function initial_para(p$) As String
Dim i%
For i% = 1 To Len(p$)
 If Mid$(p$, i%, i) = "+" Then
 initial_para = Mid$(p$, 1, i% - 1) + "#" + Mid$(p$, i% + 1, Len(p$) - i%)
 ElseIf Mid$(p$, i%, i) = "-" Then
 initial_para = Mid$(p$, 1, i% - 1) + "@" + Mid$(p$, i% + 1, Len(p$) - i%)
 ElseIf Mid$(p$, i%, i) = "/" Then
 initial_para = Mid$(p$, 1, i% - 1) + "&" + Mid$(p$, i% + 1, Len(p$) - i%)
 End If
Next i%
If Mid$(p$, Len(p$), 1) <> ")" Then
For i% = Len(p$) To 1 Step -1
 If Mid$(p$, i%, 1) = "(" Then
  p$ = Mid$(p$, 1, i%) + add_brace( _
    Mid$(p$, i% + 1, Len(p$) - i%), "")
   Exit Function
 End If
  p$ = add_brace(p$, "")
Next i%
End If
End Function


Private Function combine_divide_string(ByVal p1$, ByVal p2$) As String
Dim ts1$
Dim ts2$
If InStr(1, p1$, "F", 0) > 0 Or InStr(1, p2$, "F", 0) > 0 Then
 combine_divide_string = "F"
  Exit Function
ElseIf p2$ = "-1" Or p2$ = "@1" Then
p1$ = time_string(p1$, "-1", True, False)
p2$ = "1"
End If
Call string_type(p2$, "", "", ts1$, ts2$)
If ts2$ = "" And ts1$ = "1" Then
p2$ = add_brace(p2$, "para")
Else
p2$ = add_brace(p2$, "string")
End If
Call string_type(p1$, "", "", ts1$, ts2$)
If ts2$ = "" And ts1$ = "1" Then
p1$ = add_brace(p1$, "para")
Else
p1$ = add_brace(p1$, "string")
End If
If p2$ = "1" Then
combine_divide_string = p1$
Else
combine_divide_string = p1$ + "/" + p2$
End If
End Function

Public Function val_(ByVal s$) As Variant
If Mid$(s$, 1, 1) = "@" Then
s$ = "-" + Mid$(s$, 2, Len(s$) - 1)
End If
val_ = val(value_string(s$))
End Function
Public Function val0(ByVal s$, v As Variant) As Boolean
Dim ty As Byte
Dim tv(2) As Variant
Dim ts(3) As String
ty = string_type(s$, ts(0), ts(1), ts(2), ts(3))
If ty = 3 Then
 If val0(ts(1), tv(0)) And val0(ts(2), tv(1)) Then
 v = tv(0) / tv(1)
 val0 = True
 Else
 val0 = False
 End If
Else
 If ts(3) = "" Then
  If val0_para(ts(0), tv(0), 1) And val0_item(ts(2), tv(1)) Then
   v = tv(0) * tv(1)
    val0 = True
  Else
   val0 = False
  End If
 Else
 If val0(ts(0), tv(0)) And val0(ts(3), tv(1)) Then
 v = tv(0) + tv(1)
 val0 = True
 Else
 val0 = False
 End If
  
 End If
End If
End Function

Private Function str_(v As Variant) As String
Dim s As String
Dim ts(2) As String
Dim p%
 s = str(v)
 If InStr(1, s, "E", 0) > 0 Then
 str_ = "F"
 Else
s = Trim(s)
p% = InStr(1, s, ".", 0)
If p% = 0 Then
 str_ = s
  GoTo str_mark0
ElseIf p% = 1 Then
 s = "0" + s
  p% = p% + 1
End If
If Len(s) > p% + 4 Then
 str_ = Mid$(s, 1, p% + 4)
Else
 str_ = s
End If
str_mark0:
'**********************
'If Mid$(str_, 1, 1) = "-" Then
'str_ = "@" + Mid$(str_, 2, Len(str_) - 1)
'End If
End If
End Function

Public Function simple_multi_string0(s1$, S2$, s3$, S4$, fa$, is_simple As Boolean) As Boolean
Dim tp(3) As String
Dim it(4) As String
Dim tp1(7) As String
Dim tn(4) As Long
Dim gcf As Long
Dim ls$
Dim fs$
Dim tfa$
Dim i%, n%
Dim ty As Integer
Dim tvs As v_string
On Error GoTo simple_multi_string0_error
simple_multi_string0 = True
'**************
If s1$ = "0" And S2$ = "0" And s3$ = "0" And S4$ = "0" Then
   Exit Function
ElseIf s1$ = "0" Or s1$ = "" Then
   simple_multi_string0 = simple_multi_string0(S2$, s3$, S4$, "0", fa$, is_simple)
End If
If regist_data.run_type = 1 Then
   If InStr(1, s1$, "U", 0) > 0 Or InStr(1, s1$, "V", 0) > 0 Or _
        InStr(1, S2$, "U", 0) > 0 Or InStr(1, S2$, "V", 0) > 0 Or _
         InStr(1, s3$, "U", 0) > 0 Or InStr(1, s3$, "V", 0) > 0 Or _
          InStr(1, S4$, "U", 0) > 0 Or InStr(1, S4$, "V", 0) > 0 Then
      If s1$ <> "" Then
      tp(0) = simple_v_string(s1$)
      End If
      If S2$ <> "" Then
      tp(1) = simple_v_string(S2$)
      End If
      If s3$ <> "" Then
      tp(2) = simple_v_string(s3$)
      End If
      If S4$ <> "" Then
      tp(3) = simple_v_string(S4$)
      End If
      Call simple_multi_string0(tp(0), tp(1), tp(2), tp(3), fa$, is_simple)
      If fa$ <> "1" Then
       s1$ = divide_string(s1$, fa$, True, False)
       S2$ = divide_string(S2$, fa$, True, False)
       s3$ = divide_string(s3$, fa$, True, False)
       S4$ = divide_string(S4$, fa$, True, False)
      End If
       Exit Function
   End If
End If
If fa$ = "" Then
 fa$ = "1"
End If
 fs$ = time_string(fa$, s1$, True, False)
If s1$ <> "0" And S2$ = "0" And s3$ = "0" And s3$ = "0" Then
'单项
  fa$ = fs$
    s1$ = "1"
ElseIf s1$ <> "0" And S2$ <> "0" And s3$ = "0" And S4$ = "0" Then
  tfa$ = divide_string(S2$, s1$, False, False)
   If string_type(tfa$, tp(0), tp(1), tp(2), tp(3)) = 3 Then
    S2$ = tp(1)
     s1$ = tp(2)
   Else
    s1$ = "1"
     S2$ = tfa$
   End If
         fa$ = divide_string(fs$, s1$, True, False)
Else
'***********
If string_type(s1$, "", tp(0), tp(1), "") = 3 Then
 s1$ = tp(0)
 S2$ = time_string(S2$, tp(1), False, False)
 s3$ = time_string(s3$, tp(1), False, False)
 S4$ = time_string(S4$, tp(1), False, False)
 'fa$ = divide_string(fa$, tp(1), False, False)
End If
If string_type(S2$, "", tp(0), tp(1), "") = 3 Then
 S2$ = tp(0)
 s1$ = time_string(s1$, tp(1), False, False)
 s3$ = time_string(s3$, tp(1), False, False)
 S4$ = time_string(S4$, tp(1), False, False)
 'fa$ = divide_string(fa$, tp(1), False, False)
End If
If string_type(s3$, "", tp(0), tp(1), "") = 3 Then
 s3$ = tp(0)
 S2$ = time_string(S2$, tp(1), False, False)
 s1$ = time_string(s1$, tp(1), False, False)
 S4$ = time_string(S4$, tp(1), False, False)
 'fa$ = divide_string(fa$, tp(1), False, False)
End If
If string_type(S4$, "", tp(0), tp(1), "") = 3 Then
 S4$ = tp(0)
 S2$ = time_string(S2$, tp(1), False, False)
 s3$ = time_string(s3$, tp(1), False, False)
 s1$ = time_string(s1$, tp(1), False, False)
 'fa$ = divide_string(fa$, tp(1), False, False)
End If
If s1$ = "1" Or s1$ = "0" Then
 fa$ = fs$
 Exit Function
ElseIf Mid$(s1$, 1, 1) = "-" Then
s1$ = time_string("-1", s1$, is_simple, False)
S2$ = time_string("-1", S2$, is_simple, False)
s3$ = time_string("-1", s3$, is_simple, False)
S4$ = time_string("-1", S4$, is_simple, False)
fa$ = divide_string(fs$, s1$, is_simple, False)
Call simple_multi_string0(s1$, S2$, s3$, S4$, fa$, is_simple)
Else
tp(0) = s1$
 tp(1) = S2$
  tp(2) = s3$
   tp(3) = S4$
If S2$ = "0" Then
     fa$ = divide_string(fs$, s1$, is_simple, False)
       s1$ = "1"
        Exit Function
Else
'******************
For i% = 0 To 3
If tp(i%) <> "" Or tp(i%) <> "0" Then
 ty = string_type(tp(i%), "", tp(i%), tp1(i%), ls$)
  If ty = 3 Then
   tn(i%) = val(tp1(i%))
    If tn(i%) = 0 Then
     Exit Function
    End If
    ty = string_type(tp(i%), "", tp(i%), it(i%), ls$)
     If ty <> 0 Or ls$ <> "" Then
      Exit Function
     End If
   ElseIf ty = 0 And ls$ = "" Then
    tn(i%) = 1
     it(i%) = tp1(i%)
   Else
    Exit Function
   End If
 Else
  tn(i%) = 1
   it(i%) = "0"
 End If
Next i%
gcf = tn(0)
'*************
For i% = 1 To 3
 gcf = gcf * tn(i%) / l_gcd(gcf, tn(i%))
Next i%
For i% = 0 To 3
 tn(i%) = gcf / tn(i%)
  tp(i%) = time_para(tp(i%), str_(tn(i%)), is_simple, False)
Next i%
Call simple_multi_para0(tp(0), tp(1), tp(2), tp(3), tfa$)
Call simple_multi_item(it(0), it(1), it(2), it(3), "", "", "", "", "", 4, it(4))
If s1$ <> "0" And s1$ <> "" Then
 s1$ = time_string(tp(0), it(0), True, False)
End If
If S2$ <> "0" And S2$ <> "" Then
 S2$ = time_string(tp(1), it(1), True, False)
End If
If s3$ <> "0" And s3$ <> "" Then
 s3$ = time_string(tp(2), it(2), True, False)
End If
If S4$ <> "0" And S4$ <> "" Then
 S4$ = time_string(tp(3), it(3), True, False)
End If
fa$ = divide_string(fs$, s1$, is_simple, False)
End If
End If
End If
Exit Function
simple_multi_string0_error:
simple_multi_string0 = False
End Function

Public Function rational_string(s1$, S2$, s3$) As Boolean '有理化
Dim s(8) As String
Call string_type(s1$, s(7), s(0), s(1), s(2))
If s(2) = "" Then
 If rational_para(s(0), s(3), s(4)) Then
 Call rational_item(s(1), s(5), s(6))
  S2$ = time_string(s(3), s(5), True, False)
   s3$ = time_string(s(4), s(6), True, False)
      If Mid$(s3$, 1, 1) = "-" Then
       s3$ = time_string(s3$, "-1", True, False)
        S2$ = time_string(S2$, "-1", True, False)
      End If
    rational_string = True
 Else
  rational_string = False
 End If
Else '
 Call string_type(s(2), s(8), s(3), s(4), s(2))
 If s(2) = "" Then
  s(0) = ""
  Call simple_multi_string0(s(7), s(8), "", "", s(0), False)
   Call read_para_from_item(s(7), s(1), s(2))
    Call read_para_from_item(s(8), s(3), s(4))
Call rational_string(s(0), s(5), s(6))
If (InStr(1, s(2), "'", 0) > 0 Or InStr(1, s(4), "'", 0) > 0) And InStr(1, s(0), "'", 0) = 0 Then '根式和(差)
  S2$ = add_string(s(7), time_string(s(8), "-1", False, False), False, False)
    s3$ = time_string(s(7), s(7), False, False)
     s3$ = minus_string(s3$, time_string(s(8), s(8), False, False), True, False)
      s3$ = time_string(s3$, s(0), True, False)
      If Mid$(s3$, 1, 1) = "-" Then
       s3$ = time_string(s3$, "-1", True, False)
        S2$ = time_string(S2$, "-1", True, False)
      End If
   rational_string = True
ElseIf InStr(1, s(0), "'", 0) > 0 Then
  If s(5) <> "1" Then
  S2$ = s(5)
   s3$ = time_string(s(6), add_string(s(7), s(8), False, False), False, False)
      If Mid$(s3$, 1, 1) = "-" Then
       s3$ = time_string(s3$, "-1", True, False)
        S2$ = time_string(S2$, "-1", True, False)
      End If
  rational_string = True
  End If
End If
End If
End If
End Function
Private Function read_an_element_from_string(s$) As String
Dim ty As Byte
Dim ts(2) As String
Dim i%
ty = string_type(s$, "", ts(0), ts(1), ts(2))
If ty = 0 Then
If ts(1) <> "1" Then
 read_an_element_from_string = Mid$(ts(1), 1, 1)
  Exit Function
ElseIf ts(2) <> "0" Then
read_an_element_from_string = read_an_element_from_string(ts(2))
End If
End If
End Function
Private Function read_a_prime_number_from_para(p$) As Long
Dim i%
Dim tp(2) As String
Dim ty As Integer
ty = para_type(p$, "", tp(0), tp(1), tp(2))
If ty = 1 Then
read_a_prime_number_from_para = read_a_prime_number_from_int(val(tp(1)))
ElseIf ty = 0 Then
 If tp(1) <> "1" Then
read_a_prime_number_from_para = read_a_prime_number_from_int(val(tp(1)))
 Else
read_a_prime_number_from_para = read_a_prime_number_from_para(tp(2))
 End If
End If
End Function
Private Function read_a_prime_number_from_int(m&) As Long
Dim j&
Dim k&
k& = sqr(m&)
For j& = k& To 1 Step -1
If m& Mod j& = 0 Then
 If is_prime(j&) = False Then
  read_a_prime_number_from_int = j&
   Exit Function
 End If
End If
Next j&
read_a_prime_number_from_int = m&
End Function
Private Function is_prime(n&) As Boolean
Dim i&, k&
k& = sqr(n&)
For i& = 2 To k&
 If n& Mod i& = 0 Then
  is_prime = False
   Exit Function
 End If
Next i&
  is_prime = True
End Function
Private Sub trans_para_to_polynorm(s$, n&, PI As para_item_type)
Dim i%, j%, n1%, n2%, m1%, m2%
Dim p&
Dim t_pi As para_item_type
PI.last_it = 0
Call read_para(s$, t_pi)
n1% = 0
n2% = 0
m1% = 1
For i% = 0 To t_pi.last_it - 1
Call read_a_prime_from_item(t_pi.it(i%), n&, m2%, t_pi.it(i%))
If m1% - m2% = 1 Then
  ReDim Preserve PI.it(PI.last_it) As String
   ReDim Preserve PI.pA(PI.last_it) As String
 n2% = i% - 1
    p& = 1
     For j% = 1 To m1%
      p& = p& * n&
     Next j%
      PI.it(PI.last_it) = str_(p&)
  Call para_from_part_pi(t_pi, n1%, n2%, PI.pA(PI.last_it))
   PI.last_it = PI.last_it + 1
    m1% = m1% - 1
End If
Next i%
n2% = i% - 1
If n2% > n1% Then
ReDim Preserve PI.it(PI.last_it) As String
ReDim Preserve PI.pA(PI.last_it) As String
    p& = 1
     For j% = 1 To m1%
      p& = p& * n&
     Next j%
PI.it(PI.last_it) = "1"
Call para_from_part_pi(t_pi, n1%, n2%, PI.pA(PI.last_it))
PI.last_it = PI.last_it + 1
End If
End Sub

Private Sub trans_string_to_polymo(s$, ch As String, PI As para_item_type)
Dim i%, j%, k%, n1%, n2%, m1%, m2%
Dim t_pi As para_item_type
PI.last_it = 0
Call read_string(s$, t_pi)
n1% = 0
n2% = 0
m1% = 0
For i% = 0 To t_pi.last_it - 1
Call read_an_element_from_item(t_pi.it(i%), ch, m2%, t_pi.it(i%))
If m2% - m1% = 1 Then
 n2% = i% - 1
  ReDim Preserve PI.it(PI.last_it) As String
   ReDim Preserve PI.pA(PI.last_it) As String
    PI.it(PI.last_it) = ""
     For j% = 1 To m1%
      PI.it(PI.last_it) = PI.it(PI.last_it) + ch
     Next j%
  Call string_from_part_pi(t_pi, n1%, n2%, PI.pA(PI.last_it))
   PI.last_it = PI.last_it + 1
    m1% = m1% + 1
  n1% = i%
ElseIf m2% - m1% > 1 Then
 For j% = m1% To m2% - 1
  ReDim Preserve PI.it(PI.last_it) As String
   ReDim Preserve PI.pA(PI.last_it) As String
    PI.it(PI.last_it) = ""
     For k% = 1 To j%
      PI.it(PI.last_it) = PI.it(PI.last_it) + ch
     Next k%
   PI.pA(PI.last_it) = "0"
 Next j%
 m1% = m2%
 n1% = i%
End If
Next i%
n2% = i% - 1
If n2% > n1% Then
ReDim Preserve PI.it(PI.last_it) As String
ReDim Preserve PI.pA(PI.last_it) As String
PI.it(PI.last_it) = ""
For j% = 1 To m1%
PI.it(PI.last_it) = PI.it(PI.last_it) + ch
Next j%
Call string_from_part_pi(t_pi, n1%, n2%, PI.pA(PI.last_it))
PI.last_it = PI.last_it + 1
End If
End Sub
Private Sub read_a_prime_from_item(ByVal it0$, p&, n%, it1$)
Dim m&
Dim th$
m& = val(it0$)
n% = 0
Do While m& > p&
If m& Mod p& = 0 Then
 m& = m& / p&
  n% = n% + 1
Else
it1$ = str_(m&)
Exit Sub
End If
Loop
End Sub
Private Sub read_an_element_from_item(ByVal it0$, ch$, n%, it1$)
Dim i%
Dim th$
it1$ = ""
n% = 0
For i% = 1 To Len(it0$)
th$ = Mid$(it0$, i%, 1)
If th$ = ch$ Then
n% = n% + 1
Else
it1$ = it1$ + th$
End If
Next i%
End Sub

Private Sub string_from_part_pi(PI As para_item_type, n1%, n2%, pA$)
Dim i%
pA$ = ""
For i% = n1% To n2%
pA$ = combine_para_or_string_for_add(pA$, _
        time_string(PI.pA(i%), PI.it(i%), True, False), "string")
Next i%
End Sub
Private Sub para_from_part_pi(PI As para_item_type, n1%, n2%, pA$)
Dim i%
pA$ = ""
For i% = n1% To n2%
pA$ = combine_para_or_string_for_add(pA$, _
        combine_para_for_item(PI.pA(i%), PI.it(i%)), "para")
Next i%
End Sub
Public Sub pseudodivide_for_para(ByVal p1$, ByVal p2$, q$, f1$, r$, is_simple As Boolean)
Dim i%, j%, l%
Dim m_l%
Dim k&, k1&, k2&, k3&
Dim pi1 As para_item_type
Dim pi2 As para_item_type
Dim tpi As para_item_type
Dim tf() As String
Dim f(1) As String
Dim tq() As String
Dim q_(1) As String
Dim tr$
Dim tp(1) As String
Dim last_it As Integer
k1& = read_a_prime_number_from_para(p1$)
k2& = read_a_prime_number_from_para(p2$)
k& = max(k1&, k2&)
Call trans_para_to_polynorm(p1$, k&, pi1)
Call trans_para_to_polynorm(p2$, k&, pi2)
f1$ = "1"
For i% = 0 To pi1.last_it - 1
If pi1.it(i%) < pi2.it(0) Then
  Call para_from_part_pi(pi1, i%, pi1.last_it - 1, r$)
   Exit Sub
Else
            k1& = val(pi1.it(i%))
             k2& = val(pi2.it(0))
              k& = k1& / k2&
     If InStr(1, pi1.pA(0), "'", 0) = 0 And _
            InStr(1, pi2.pA(0), "'", 0) = 0 Then
         k1& = val_(pi1.pA(i%))
          k2& = val_(pi2.pA(0))
       Call simple_two_long(k1&, k2&, k&)
         f(0) = str_(k2&)
          f(1) = str_(k1&)
      For j% = i% To pi1.last_it - 1
        pi1.pA(j%) = time_para(pi1.pA(j%), f(0), is_simple, False)
      Next j%
      For j% = 0 To pi2.last_it - 1
        pi1.pA(i% + j%) = minus_para(pi1.pA(i% + j%), _
               time_para(pi2.pA(j%), f(1), False, False), is_simple, False)
      Next j%
     f1$ = time_para(f1$, f(0), is_simple, False)
     q$ = add_para(time_para(f(0), q$, False, False), _
            combine_para_for_item(tf(1), str_(k&)), is_simple, False)
     Else
      tp(0) = pi1.pA(j%)
       tp(1) = pi2.pA(0)
        l = 0
        q_(0) = ""
        q_(1) = ""
        last_it = 0
        Erase tf
        Erase tq
        Do
         ReDim Preserve tf(last_it) As String
          ReDim Preserve tq(last_it) As String
           m_l% = (l% + 1) Mod 2
           Call pseudodivide_for_para(tp(l%), tp(m_l%), _
               tq(last_it), tf(last_it), tr$, is_simple)
         tp(l%) = tp(m_l%)
          tp(m_l%) = tr$
          last_it = last_it + 1
        Loop While tr$ <> "" Or tr$ <> "0"
     tr$ = tp((l% + 1) Mod 2)  'pseudo_gcd
 For j% = last_it - 2 To 0 Step -1
  tf(j%) = time_para(tf(j%), tq(j% + 1), False, False)
   tq(j%) = add_para(tf(j% + 1), time_para(tq(j%), tq(j% + 1), False, False), _
               is_simple, False)
 Next j%
 For j% = i% To pi1.last_it - 1
  pi1.pA(j%) = time_para(pi1.pA(j%), tf(0), is_simple, False)
 Next j%
   For j% = 0 To pi2.last_it - 1
        pi1.pA(i% + j%) = minus_para(pi1.pA(i% + j%), _
               time_para(pi2.pA(j%), tq(0), False, False), is_simple, False)
   Next j%
   f1$ = time_para(f1$, tf(0), is_simple, False)
   q$ = add_para(time_para(tf(0), q$, False, False), _
           time_para(tq(0), str_(k&), False, False), is_simple, False)
  End If
End If
Next i%
End Sub
Public Sub pseudodivide_for_string(ByVal s1$, ByVal S2$, q$, f1$, r$, is_simple As Boolean)
Dim i%, j%, l%
Dim m_l%
Dim k$, k1$, k2$, k3$
Dim pi1 As para_item_type
Dim pi2 As para_item_type
Dim tpi As para_item_type
Dim tf() As String
Dim f(1) As String
Dim tq() As String
Dim q_(1) As String
Dim tr$
Dim tp(1) As String
Dim last_it As Integer
k1$ = read_an_element_from_string(s1$)
k2$ = read_an_element_from_string(S2$)
If k1$ > k2$ Then
k1$ = k2$
End If
Call trans_string_to_polymo(s1$, k1$, pi1)
Call trans_string_to_polymo(S2$, k1$, pi2)
f1$ = "1"
For i% = pi1.last_it - 1 To 0 Step -1
If pi1.it(i%) < pi2.it(pi2.last_it - 1) Then
  Call string_from_part_pi(pi1, 0, i%, r$)
   Exit Sub
Else
            k1$ = pi1.it(i%)
             k2$ = pi2.it(pi2.last_it - 1)
            Call divide_item(k1$, k2$, k$, "")
     If read_an_element_from_string(pi1.pA(i%)) = "" And _
         read_an_element_from_string(pi2.pA(pi2.last_it - 1)) = "" Then
         k1$ = pi1.pA(i%)
          k2$ = pi2.pA(pi2.last_it - 1)
        Call simple_two_para(k1$, k2$, k$)
         f(0) = k2$
          f(1) = k1$
      pi1.pA(i%) = "0"
       pi1.it(i%) = ""
      For j% = i% - 1 To 0 Step -1
        pi1.pA(j%) = time_para(pi1.pA(j%), f(0), True, False)
      Next j%
      For j% = pi2.last_it - 2 To 0 Step -1
        pi1.pA(i% - j% + pi2.last_it - 1) = _
           minus_string(pi1.pA(i% - j% + pi2.last_it - 1), _
               time_string(pi2.pA(j%), f(1), False, False), is_simple, False)
      Next j%
     f1$ = time_string(f1$, f(0), is_simple, False)
     q$ = add_string(time_para(f(0), q$, False, False), _
            time_string(tf(1), k$, False, False), is_simple, False)
     Else
      tp(0) = pi1.pA(j%)
       tp(1) = pi2.pA(0)
        l = 0
        q_(0) = ""
        q_(1) = ""
        last_it = 0
        Erase tf
        Erase tq
        Do
         ReDim Preserve tf(last_it) As String
          ReDim Preserve tq(last_it) As String
           m_l% = (l% + 1) Mod 2
           Call pseudodivide_for_string(tp(l%), tp(m_l%), _
               tq(last_it), tf(last_it), tr$, is_simple)
         tp(l%) = tp(m_l%)
          tp(m_l%) = tr$
          last_it = last_it + 1
        Loop While tr$ <> "" Or tr$ <> "0"
     tr$ = tp(m_l%)  'pseudo_gcd
 For j% = last_it - 2 To 0 Step -1
  tf(j%) = time_para(tf(j%), tq(j% + 1), False, False)
   tq(j%) = add_para(tf(j% + 1), time_para(tq(j%), tq(j% + 1), False, False), _
                 is_simple, False)
 Next j%
 pi1.pA(i%) = "0"
 pi1.it(i%) = ""
 For j% = i% - 1 To 0 Step -1
  pi1.pA(j%) = time_string(pi1.pA(j%), tf(0), is_simple, False)
 Next j%
   For j% = pi2.last_it - 2 To 0
        pi1.pA(i% - j% + pi2.last_it - 1) = minus_para(pi1.pA(i% - j% + pi2.last_it - 1), _
               time_para(pi2.pA(j%), tq(0), False, False), is_simple, False)
   Next j%
   f1$ = time_para(f1$, tf(0), False, False)
   q$ = add_para(time_para(tf(0), q$, False, False), _
           time_string(tq(0), k$, False, False), is_simple, False)
  End If
End If
Next i%
End Sub

Private Sub simple_two_para(p1$, p2$, p3$)
Call simple_multi_para(p1$, p2$, "", "", "", "", "", "", "", 2, p3$, False, False)
End Sub
Private Sub simple_two_string(s1$, S2$, s3$)
Dim i%
Dim ts(1) As String
Dim f() As String
Dim q() As String
Dim r$
Dim last_it As Integer
ts(0) = s1$
ts(1) = S2$
Do
ReDim Preserve f(last_it) As String
ReDim Preserve q(last_it) As String
Call pseudodivide_for_para(ts(0), ts(1), f(last_it), q(last_it), r$, True)
If r$ <> "" Or r$ <> "0" Then
ts(0) = ts(1)
ts(1) = r$
End If
last_it = last_it + 1
Loop While r$ <> "" Or r$ <> "0"
s3$ = ts(1)
For i% = last_it - 2 To 0 Step -1
f(i%) = time_para(f(i%), q(i% + 1), True, False)
q(i%) = add_para(f(i% + 1), time_para(q(i%), q(i% + 1), False, False), True, False)
Next i%
s1$ = q(0)
S2$ = f(0)
End Sub
Public Sub set_squre_root_string(ByVal s1$, S2$)
Dim i%, k%
Dim ts(2) As String
Dim ch As String
If InStr(1, s1$, "F", 0) > 0 Then
 S2$ = "F"
  Exit Sub
End If
If Mid$(s1$, 1, 1) = "(" Then
 k% = 1
For i% = 2 To Len(s1$)
 ch = Mid$(s1$, i%, 1)
 If ch = "(" Then
  k% = k + 1
 ElseIf ch = ")" Then
  k% = k% - 1
   If i% = Len(s1$) Then
    Call set_squre_root_string(Mid$(s1$, 2, Len(s1$) - 2), S2$)
     Exit Sub
   End If
 End If
Next i%
End If
If s1$ = "1" Then
S2$ = s1$
ElseIf Len(s1$) = 1 And s1$ >= "a" And s1$ <= "z" Then
S2$ = "'" + s1$
Else
 If InStr(1, s1$, "[", 0) > 0 Then 'ch <> "'" And Asc(ch) < 0 And ch <> "\" Then
  S2$ = "F"
   Exit Sub
 Else
  Call string_type(s1$, ts(1), "", "", ts(0))
 End If
 S2$ = "[" + Trim(str(set_number_string(s1$, 0))) + "]"
 End If
End Sub
Public Function set_squre_root_string_(ByVal s1$) As String
Call set_squre_root_string(ByVal s1$, set_squre_root_string_)
End Function
Public Function solut_2order_equation(ByVal A$, ByVal b$, _
   ByVal c$, s1$, S2$, cal_float As Boolean) As Boolean
Dim delta$
delta$ = minus_string(time_string(b$, b$, False, cal_float), _
         time_string(4, time_string(A$, c$, False, cal_float), _
          False, cal_float), False, cal_float)
If is_less_than(delta$, "0") Then
 solut_2order_equation = False
  Exit Function
Else
 solut_2order_equation = True
 delta$ = sqr_string(delta$, False, cal_float)
 If Mid$(A$, 1, 1) = "-" Then
  A$ = time_string("-2", A$, False, cal_float)
 Else
 A$ = time_string(2, A$, False, cal_float)
     b$ = time_string("-1", b$, False, cal_float)
 End If
 s1$ = divide_string(add_string(b$, delta$, False, cal_float), A$, True, cal_float)
 S2$ = divide_string(minus_string(b$, delta$, False, cal_float), A$, True, cal_float)
End If
End Function

Public Sub exchange_two_integer(n1%, n2%)
'交换两个基本点整数
Dim i%
 i% = n1%
  n1% = n2%
   n2% = i%
End Sub
Public Sub exchange_two_long_integer(n1 As Long, n2 As Long) '交换两个长整数
Dim i As Long
 i = n1
  n1 = n2
   n2 = i
End Sub

Public Sub exchange_two__string(s1$, S2$) '交换两个基本点字符
Dim ts As String
ts = s1$
 s1$ = S2$
  S2$ = ts
End Sub
Public Function is_simple(ByVal s As String, ty As Byte) As Boolean
Dim i%, j%, k%
If Mid$(s, Len(s), 1) <> ")" Then
 is_simple = False
Else
If ty = 0 Then 'string
If InStr(1, s, "/", 0) > 0 Then
 is_simple = True
Else
 i% = InStr(1, s, "(", 0)
 j% = InStr(2, s, "+", 0)
 k% = InStr(2, s, "-", 0)
 If k% = 0 And j% = 0 Then
  is_simple = True
 Else
  If j% < k% And j% > 0 Then
   k% = j%
  End If
   If i% > k% Then
    is_simple = False
   Else
    If right_brace(s, i% + 1) = Len(s) Then
     is_simple = True
    Else
     is_simple = False
    End If
   End If
 End If
End If
Else
  If InStr(1, s, "&", 0) > 0 Then
   is_simple = True
  Else
   i% = InStr(1, s, "(", 0)
   j% = InStr(2, s, "#", 0)
   k% = InStr(2, s, "@", 0)
 If k% = 0 And j% = 0 Then
  is_simple = True
 Else
  If j% < k% And j% > 0 Then
   k% = j%
  End If
   If i% > k% Then
    is_simple = False
   Else
    If right_brace(s, i% + 1) = Len(s) Then
     is_simple = True
    Else
     is_simple = False
    End If
   End If
 End If
End If
End If
End If
End Function

Public Function max_for_byte(n1 As Byte, n2 As Byte) As Byte
If n1 >= n2 Then
 max_for_byte = n1
Else
 max_for_byte = n2
End If
End Function
Public Function min_for_byte(n1 As Byte, n2 As Byte) As Byte
If n1 >= n2 Then
 min_for_byte = n2
Else
 min_for_byte = n1
End If
End Function


Public Function solve_first_order_equation(ByVal s1$, ByVal S2$, unknown_element As String) As String
Dim ts As String
Dim ch$
Dim para$
If s1$ = "" Or InStr(1, s1$, "F", 0) > 0 Or S2$ = "" Or _
       InStr(1, S2$, "F", 0) > 0 Or unknown_element = "" Or _
        InStr(1, unknown_element, "F", 0) > 0 Then
 solve_first_order_equation = "F"
 Exit Function
End If
para$ = "0"
ch$ = "0"
ts = minus_string(s1$, S2$, True, False)
Call read_para_and_const_from_first_order_equation(ts, unknown_element, para$, ch$)
 solve_first_order_equation = divide_string(time_string("-1", ch$, False, False), para$, True, False)
End Function
Private Function divide_string_for_one_item(ByVal s1 As String, ByVal S2 As String, _
       q As String) As Boolean
Dim t_q As String
Dim ts(1) As String
Dim pA(1) As String
Dim it(1) As String
Dim i%
Dim ty As Byte
Dim p(1) As String
Dim iem(1) As String
On Error GoTo divide_string_for_one_item_error
 Call read_para_from_item(s1, pA(0), it(0))
 Call read_para_from_item(S2, pA(1), it(1))
pA(0) = divide_para(pA(0), pA(1), True, False)
Call simple_multi_item(it(1), it(0), "", "", "", "", "", "", "", 2, "")
t_q = divide_string(s1, S2, True, False)
If it(1) = "1" Then
 q = time_string(pA(0), it(0), True, False)
  If q = "" Then
   q = "0"
  End If
Else
  q = "0"
   divide_string_for_one_item = True
End If
Exit Function
divide_string_for_one_item_error:
End Function

Private Function divide_string_with_remainder0(ByVal s1 As String, ByVal S2 As String, _
      q As String, r As String) As Boolean
Dim ty(1) As Byte
Dim ts(4) As String
Dim tq As String
Dim tr As String
On Error GoTo divide_string_with_remainder_error
ty(0) = string_type(s1, ts(0), "", "", ts(2))
ty(1) = string_type(S2, ts(1), "", "", "")
If ty(0) = 3 Or ty(1) = 3 Then
GoTo divide_string_with_remainder_error
Else
 If divide_string_for_one_item(ts(0), ts(1), tq) Then
   q = tq
  If q = "0" Then
   r = s1
  Else
    If ts(2) = "" Then
       ts(2) = "0"
    End If
    ts(1) = time_string(S2, tq, False, False)
    Call string_type(ts(1), "", "", "", ts(3))
    If ts(3) = "" Then
     ts(3) = "0"
    End If
    r = minus_string(ts(2), ts(3), True, False)
  End If
   divide_string_with_remainder0 = True
 End If
End If
Exit Function
divide_string_with_remainder_error:
divide_string_with_remainder0 = False
End Function
Private Function divide_string_with_remainder(ByVal s1 As String, ByVal S2 As String, _
      q As String, r As String) As Boolean
Dim ts() As String
Dim tr(1) As String
Dim last_ts As Integer
Dim i%
On Error GoTo divide_string_with_remainder_error
q = "0"
tr(0) = s1
divide_string_with_remainder_back:
last_ts = last_ts + 1
ReDim Preserve ts(last_ts) As String
If divide_string_with_remainder0(tr(0), S2, ts(last_ts), tr(1)) Then
 If ts(last_ts) <> "0" Then '商不等于零
   If tr(1) = "0" Then '余为零
      q = add_string(q, ts(last_ts), True, False)
        r = "0"
   Else '余不为零
    tr(0) = tr(1)
      GoTo divide_string_with_remainder_back
   End If
 Else
   r = tr(1)
 End If
divide_string_with_remainder = True
Else
divide_string_with_remainder = False
End If
Exit Function
divide_string_with_remainder_error:
End Function
Public Function division_string(ByVal s1 As String, ByVal S2 As String, gcd As String, _
          q1 As String, q2 As String) As Boolean '辗转相除
Dim ts(1) As String
Dim tq As String
Dim tr As String
Dim ty(1) As Integer
Dim ty_(1) As Integer
ts(0) = s1
ts(1) = S2
ty(0) = InStr(1, s1, "x", 0)
ty(1) = InStr(1, S2, "x", 0)
ty_(0) = InStr(1, s1, "[", 0)
ty_(1) = InStr(1, S2, "[", 0)
If (ty(0) > 0 And ty(1) > 0) And ty_(0) = 0 And ty_(1) = 0 Then
division_string_back:
If divide_string_with_remainder(ts(0), ts(1), tq, tr) Then
 If tr <> "0" Then
  ts(0) = ts(1)
   ts(1) = tr
    GoTo division_string_back:
 Else
  If tq <> "0" Then
  gcd = ts(1)
   Call divide_string_with_remainder(s1, gcd, q1, "")
    Call divide_string_with_remainder(S2, gcd, q2, "")
     division_string = True
  Else
    tr = ts(0)
    ts(0) = ts(1)
    ts(1) = tr
    GoTo division_string_back:
  End If
 End If
End If
Else
'If simple_multi_string0(s1, S2, "", "", gcd, True) Then
' q1 = s1
' q2 = S2
'Else
' gcd = "1"
' q1 = s1
' q2 = S2
'End If
End If
End Function

Public Function subs_value_for_string(ByVal s$, ByVal s1$, ByVal v As String) As String '替换将S$中的S$替换为V
Dim ts(4) As String
Dim ty As Integer
Dim tn%
If is_contain_x(s$, s1$, 1) = False Then
subs_value_for_string = s$
ElseIf s$ = "x" Then
 subs_value_for_string = v
Else
ty = string_type(s$, ts(0), ts(1), ts(2), ts(3))
If ty = 3 Then
 subs_value_for_string = divide_string(subs_value_for_string(ts(1), s1$, v), _
                                       subs_value_for_string(ts(2), s1$, v), True, False)
ElseIf ty = 0 Then
  subs_value_for_string = time_string(subs_value_for_para(ts(1), s1$, v), _
                                       subs_value_for_item(ts(2), s1$, v), True, False)
  If ts(3) <> "" Then
 
 subs_value_for_string = add_string(subs_value_for_string, _
                                    subs_value_for_string(ts(3), s1$, v), True, False)
  End If
 End If
End If
End Function

Public Function subs_value_for_para(ByVal p$, ByVal s$, ByVal v As String) As String
Dim tp(3) As String
Dim ty As Integer
Dim i%
Dim it As para_item_type
If InStr(1, p$, s$, 0) = 0 Then
 subs_value_for_para = p$
Else
ty = para_type(p$, "", tp(0), tp(1), tp(2))
If ty = 0 Then
 If tp(2) = "" Then
  Call read_para(p$, it)
   For i% = 1 To it.last_it
    it.pA(i%) = subs_value_for_string_(it.pA(i%), s$, v)
    it.pA(i%) = subs_value_for_string_(it.it(i%), s$, v)
   Next i%
   subs_value_for_para = para_from_para_item(it)
 Else
  tp(3) = combine_para_for_item(tp(0), tp(1))
  subs_value_for_para = add_para(subs_value_for_para(tp(3), s$, v), _
      subs_value_for_para(tp(2), s$, v), True, False)
 End If
ElseIf ty = 1 Then
 subs_value_for_para = time_string(subs_value_for_para(tp(0), s$, v), _
       sqr_para(subs_value_for_para(tp(1), s$, v), "", "", False), True, False)
ElseIf ty = 3 Then
 subs_value_for_para = divide_para(subs_value_for_para(tp(0), s$, v), _
      subs_value_for_para(tp(1), s$, v), True, False)
Else
subs_value_for_para = p$
End If
End If
End Function

Public Function subs_value_for_item(ByVal item0$, ByVal s$, ByVal v As String) As String
 Dim tn%
 Dim tv$
 If is_contain_x(item0$, s$, 1) = False Then
  subs_value_for_item = item0$
 Else
 tn% = InStr(1, item0, "[", 0)
 If tn% > 0 Then 'Asc(item0$) < 0 And item0 <> "'" And item0 <> "\" And Len(item0$) = 1 Then
 'tn% = from_char_to_no(item0$)
  subs_value_for_item = sqr_string(subs_value_for_string(read_sqr_from_string(item0, 0, tv$), s$, v), True, False)
 Else
  subs_value_for_item = subs_value_for_string_(item0$, s$, v)
 End If
 End If
End Function

Public Function subs_value_for_string_(ByVal s$, ByVal s1$, ByVal v As String) As String
Dim ts(1) As String
Dim k%
k% = InStr(1, s$, s1$, 0)
If k% = 0 Then
 subs_value_for_string_ = s$
Else
 If s$ = s1$ Then
  subs_value_for_string_ = v
 Else
 If k% = 1 Then
 ts(0) = v
 ElseIf k% = 2 And Mid$(s$, 1, 1) = "-1" Then
 ts(0) = "-1"
 Else
 ts(0) = Mid$(s$, 1, k% - 1)
 End If
 ts(1) = Mid$(s$, k% + 1, Len(s$) - k%)
 subs_value_for_string_ = subs_value_for_string_( _
      time_para(ts(0), time_para(v, ts(1), True, False), True, False), _
        s1$, v)
 End If
End If
End Function


Public Function read_order_for_item(ByVal it As String, para As String) As Integer
Dim k%
Dim ts$
para = "1"
read_order_for_item = "0"
read_order_for_item_back:
k% = InStr(1, it, "x", 0)
If k% > 0 Then
 If k% = 1 Then
  ts$ = "1"
 ElseIf k% = 2 And Mid$(it, 1, 1) = "-1" Then
  ts$ = "-1"
 Else
  ts$ = Mid$(it, 1, k% - 1)
 End If
 para = time_string(para, ts$, True, False)
 read_order_for_item = read_order_for_item + 1
  it = Mid$(it, k% + 1, Len(it) - k%)
   GoTo read_order_for_item_back
Else
End If
End Function
Public Function read_para_from_string_for_ietm(ByVal eq$, ByVal item$, re_eq$) As String
'从eq$读出item$的系数,re_eq$余项
Dim t_item$, t_item1$, t_eq$
Dim s(5) As String
Dim ty As Byte
 t_item$ = item$
 read_para_from_string_for_ietm = "0"
 re_eq$ = "0"
 s(0) = eq$
ty = string_type(s(0), s(1), s(2), s(3), s(4))
If ty = 3 Then
 s(0) = s(2)
 s(5) = s(3)
ty = string_type(s(0), s(1), s(2), s(3), s(4))
Else
 s(5) = "1"
End If
Do
   If InStr(1, s(1), t_item$, 0) > 0 Then
    read_para_from_string_for_ietm = add_string(read_para_from_string_for_ietm, _
         divide_string(s(1), t_item$, False, False), True, False)
   Else
    re_eq$ = add_string(re_eq$, s(1), True, False)
   End If
   s(0) = s(4)
   If s(0) <> "" Then
    Call string_type(s(0), s(1), s(2), s(3), s(4))
   End If
 Loop Until s(0) = ""
 If s(5) <> "1" Then
    read_para_from_string_for_ietm = divide_string(read_para_from_string_for_ietm, s(5), True, False)
     re_eq$ = divide_string(re_eq$, s(5), True, False)
 End If
End Function
Public Function read_para_from_equation(ByVal e$, equa As Equation_data0_type) As Boolean
Dim para(6) As String
Dim ty As Integer
Dim ts(7) As String
Dim k%
para(0) = "0"
para(1) = "0"
para(2) = "0"
para(3) = "0"
para(4) = "0"
para(5) = "0"
solve_equation0_back:
If is_contain_x(e$, "x", 1) = False And is_contain_x(e$, "y", 1) = False Then
 para(0) = add_string(para(0), e, True, False) '不含未知数
  Exit Function
Else
ty = string_type(e$, ts(0), ts(1), ts(2), ts(3))
If ty = 0 Then
Do
  If InStr(1, ts(0), "xx", 0) > 0 Then ' 前部分不含
    para(0) = add_string(para(0), divide_string(ts(0), "xx", False, False), True, False)
  ElseIf InStr(1, ts(0), "yy", 0) Then    ' 前部分不含
    para(1) = add_string(para(1), divide_string(ts(0), "yy", False, False), True, False)
  ElseIf InStr(1, ts(0), "xy", 0) Then    ' 前部分不含
    para(2) = add_string(para(2), divide_string(ts(0), "xy", False, False), True, False)
  ElseIf InStr(1, ts(0), "x", 0) Then    ' 前部分不含
    para(3) = add_string(para(3), divide_string(ts(0), "x", False, False), True, False)
  ElseIf InStr(1, ts(0), "y", 0) Then    ' 前部分不含
    para(4) = add_string(para(4), divide_string(ts(0), "y", False, False), True, False)
  Else
    para(5) = add_string(para(5), ts(0), True, False)
  End If
   e$ = ts(3)
   If e$ <> "" Then
    Call string_type(e$, ts(0), ts(1), ts(2), ts(3))
   End If
 Loop Until e$ = ""
 equa.para_xx = para(0)
 equa.para_yy = para(1)
 equa.para_xy = para(2)
 equa.para_x = para(3)
 equa.para_y = para(4)
 equa.para_c = para(5)
 If equa.para_xx <> "0" Then
  Call simple_multi_string(equa.para_xx, equa.para_xy, equa.para_yy, equa.para_x, equa.para_y, _
                                equa.para_c, "0", "0", "0", 6, "0", True, False)
 ElseIf equa.para_yy <> "0" Then
  Call simple_multi_string(equa.para_yy, equa.para_x, equa.para_y, _
                                equa.para_c, "0", "0", "0", "0", "0", 4, "0", True, False)
 ElseIf equa.para_x <> "0" Then
  Call simple_multi_string(equa.para_x, equa.para_y, _
                                equa.para_c, "0", "0", "0", "0", "0", "0", 3, "0", True, False)
 ElseIf equa.para_y <> "0" Then
  Call simple_multi_string(equa.para_y, _
                                equa.para_c, "0", "0", "0", "0", "0", "0", "0", 2, "0", True, False)
 End If
read_para_from_equation = True
ElseIf ty = 3 Then
 e$ = ts(1)
  read_para_from_equation = read_para_from_equation(ByVal e$, equa)
End If
End If
'solve_equation0 = solve_equation(para(2), para(1), para(0), r1$, r2$, True)

End Function

Public Sub simple_two_string_(s1$, S2$, gcd$)
Dim i%, j%, tn%
Dim f1(3) As String
Dim f2(3) As String
If minus_string(s1$, S2$, True, False) = "0" Then
   gcd$ = s1
    s1$ = "1"
    S2$ = "1"
'ElseIf division_string(s1$, S2$, gcd$, f1(0), f1(1)) Then
'  s1$ = f1(0)
'  S2$ = f1(1)
Else
gcd = "1"
Call do_factor1(s1, f1(0), f1(1), f1(2), f1(3), 0)
Call do_factor1(S2, f2(0), f2(1), f2(2), f2(3), 0)
For i% = 0 To 3
 If f1(i%) <> "1" Then
 For j% = 0 To 3
 If f2(j%) <> "1" Then
  If j% = 0 Then
   tn% = InStr(1, f2(0), "[", 0) '读根号
   If tn% > 0 Then 'Asc(f2(0)) < 0 And f2(0) <> "'" And f2(0) <> "\" And Len(f2(0)) = 1 Then
    If f1(i%) = f2(0) Then
     gcd$ = time_string(gcd$, f2(2), True, False)
        f1(i%) = "1"
        f2(0) = "1"
    ElseIf f1(i%) = read_sqr_from_string(f2(0), 0, "") Then ' squre_root_string(tn%) Then
     f1(i%) = f2(0)
     f2(0) = "1"
    End If
   Else
    If f1(i%) = f2(0) Then
     f1(i%) = "1"
     f2(0) = "1"
    End If
   End If
  Else
  If f1(i%) = f2(j%) Then
   f1(i%) = "1"
   f2(j%) = "1"
  End If
  End If
 End If
 Next j%
 End If
Next i%
s1$ = time_string(f1(0), f1(1), False, False)
s1$ = time_string(s1$, f1(2), False, False)
s1$ = time_string(s1$, f1(3), True, False)
S2$ = time_string(f2(0), f2(1), False, False)
S2$ = time_string(S2$, f2(2), False, False)
S2$ = time_string(S2$, f2(3), True, False)
End If
End Sub
Public Function root_item(ByVal it$, root_order%, out_it$) As Boolean
Dim t_ch As String
Dim i%, j%, k%, t%, start%
If it$ = "1" Then
out_it$ = "1"
root_item = True
Else
t% = Len(it$) Mod root_order%
If t% <> 0 Then
 Exit Function
ElseIf t% Mod root_order% = 0 Then
 start% = 1
 Do While start% <= Len(it$)
  If read_same_ch_from_string(it$, start%, root_order%, t_ch, start%) Then
   out_it$ = out_it$ + t_ch
  Else
   root_item = False
    Exit Function
  End If
 Loop
 root_item = True
  Exit Function
End If
End If
End Function
Private Function read_same_ch_from_string(ByVal it$, ByVal start%, root_order%, ch As String, i%) As Boolean
Dim en%
ch = Mid$(it$, start%, 1)
en% = start% + root_order - 1
For i% = start% To en%
 If Mid$(it$, i%, 1) <> ch Then
    read_same_ch_from_string = False
      Exit Function
 End If
Next i%
    read_same_ch_from_string = True
      Exit Function
End Function
Public Function is_x_in_string(ByVal s As String) As Byte
Dim i%, is_x%
Dim tp(3) As String
Dim ty As Byte
ty = string_type(s, tp(0), tp(1), tp(2), tp(3))
If ty = 0 Then
  is_x_in_string = is_x_in_item(tp(0))
   If is_x_in_string < 1 Then
    ty = string_type(tp(3), tp(0), tp(1), tp(2), tp(3))
     is_x% = is_x_in_item(tp(0))
      If is_x% = 1 Then
       is_x_in_string = 1
      Exit Function
      End If
   End If
ElseIf ty = 3 Then
 If is_x_in_string(tp(1)) = 1 Or is_x_in_string(tp(2)) = 1 Then
  is_x_in_string = 1
 End If
End If
End Function

Public Function is_x_in_item(it$) As Byte
Dim i%, tn%
Dim ch As String
Dim it1$
Dim it2$
If InStr(1, it$, "x", 0) > 0 Then
is_x_in_item = 1
End If
For i% = 1 To Len(it$)
ch = Mid$(it$, i%, 1)
tn% = InStr(1, ch, "[", 0)
If tn% > 0 Then 'Asc(ch) < 0 And ch <> "'" And ch <> "\" Then
it1$ = Mid$(it$, i%, Len(it$) - i% + 1)
 is_x_in_item = is_x_in_item + is_x_in_item(read_sqr_from_string(it1$, 0, it2$))
  If is_x_in_item = 2 Then
   Exit Function
  Else
   is_x_in_item = is_x_in_item + is_x_in_item(it2$)
  End If
  Exit Function
End If
Next i%
End Function
Public Function val0_para(p$, vp As Variant, ty As Byte) As Boolean 'ty=0 + ty=1*
Dim i%, k%, l%, r_n%
Dim tp$
Dim th As String
Dim tvp(1) As Variant
If p$ = "" Then
 If ty = 0 Then
  vp = 0
 Else
  vp = 1
 End If
   val0_para = True
    Exit Function
End If
For i% = 1 To Len(p$)
th = Mid$(p$, i%, 1)
If th = "#" Then
tp$ = tp$ + "+"
ElseIf th = "@" Then
tp$ = tp$ + "-"
Else
tp$ = tp$ + th
End If
Next i%
k% = InStr(1, p$, "(", 0)
If k% > 0 Then
Else
k% = InStr(2, tp$, "+", 0)
l% = InStr(2, tp$, "-", 0)
If k% = 0 And l% = 0 Then
 r_n% = InStr(1, p$, "'", 0)
  If r_n% = 0 Then
   vp = val(tp$)
  Else
   If Mid$(tp$, 1, 1) = "-" And r_n% = 2 Then
    If r_n% > 2 Then
     vp = -1 * val(Mid$(tp$, 2, r_n% - 2)) * sqr(val(Mid$(tp$, r_n% + 1, Len(tp) - r_n%)))
    Else
     vp = -1 * sqr(val(Mid$(tp$, r_n% + 1, Len(tp) - r_n%)))
    End If
   ElseIf Mid$(tp$, 1, 1) = "+" And r_n% = 2 Then
        If r_n% > 2 Then
         vp = val(Mid$(tp$, 2, r_n% - 2)) * sqr(val(Mid$(tp$, r_n% + 1, Len(tp) - r_n%)))
        Else
         vp = sqr(val(Mid$(tp$, r_n% + 1, Len(tp) - r_n%)))
        End If
   Else
    If r_n% > 1 Then
     vp = val(Mid$(tp$, 1, r_n% - 1)) * sqr(val(Mid$(tp$, r_n% + 1, Len(tp) - r_n%)))
    Else
     vp = sqr(val(Mid$(tp$, r_n% + 1, Len(tp) - r_n%)))
    End If
   End If
  End If
val0_para = True
ElseIf (k% = 0 Or l% < k%) And l% > 0 Then
 If val0_para(Mid$(tp$, 1, l% - 1), tvp(0), 0) And _
       val0_para(Mid$(tp$, l%, Len(tp$) - l% + 1), tvp(1), 0) Then
        vp = tvp(0) + tvp(1)
         val0_para = True
 End If
ElseIf (l% = 0 Or k% < l%) And k% > 0 Then
 If val0_para(Mid$(tp$, 1, k% - 1), tvp(0), 0) And _
       val0_para(Mid$(tp$, k%, Len(tp$) - k% + 1), tvp(1), 0) Then
        vp = tvp(0) + tvp(1)
         val0_para = True
 End If
'ElseIf k% < l% Then
'ElseIf k% > l% Then
End If
End If
End Function
Public Function val0_item(it$, iv As Variant) As Boolean
Dim i%, k%
Dim T_IV(1) As Variant
Dim ch As String
Dim t_it(1) As String
 k% = InStr(1, it$, "'", 0)
 If k% > 0 Then
  t_it(0) = Mid$(it$, 1, k% - 1)
  t_it(1) = Mid$(it$, k% + 1, Len(it$) - k%)
  If val0_item(t_it(0), T_IV(0)) And val0_item(t_it(1), T_IV(0)) Then
   iv = T_IV(0) * sqr(T_IV(1))
    val0_item = True
  End If
 Else
  For i% = 1 To Len(it$)
   ch = Mid$(it$, i%, 1)
    If ch > "A" Then
     Exit Function
    End If
  Next i%
  iv = val(it$)
  val0_item = True
 End If
End Function
Public Function is_sqr_in_string(s As String) As Boolean
Dim i%
Dim ch As String
For i% = 1 To Len(s)
ch = Mid$(s, i%, 1)
If Asc(ch) < 0 Then
 is_sqr_in_string = True
  Exit Function
End If
Next i%
End Function

Public Function change_root_mark(ByVal s$) As String
Dim k%
k% = InStr(1, s$, "'", 0)
If k% = 0 Then
 change_root_mark = s$
Else
 change_root_mark = Mid$(s$, 1, k% - 1) + "'" + _
            change_root_mark(Mid$(s$, k + 1, Len(s$) - k%))
End If
End Function
Public Function sqr_para_for_two_root(ByVal para1$, ByVal para2$, sig As String) As String
'ty=0,初始，ty=1 调用
Dim ty_(1) As Integer
Dim s0(2) As String
Dim s1(2) As String
Dim sqr_p(1) As String
Dim r(1) As String
Dim g(2) As String
On Error GoTo sqr_para_for_two_root_error
  If solut_2order_equation("1", para1$, para2$, r(0), r(1), False) Then
      If InStr(1, r(0), "'", 0) = 0 And InStr(1, r(1), "'", 0) = 0 Then
         s1(0) = sqr_para(r(0), "", "", False)
         s1(1) = sqr_para(r(1), "", "", False)
         If sig = "-" Then
          If val(r(0)) > val(r(1)) Then
           s1(1) = time_para(s1(1), "@1", False, False)
          Else
           s1(0) = time_para(s1(0), "@1", False, False)
          End If
         End If
                   sqr_para_for_two_root = add_para(s1(0), s1(1), True, False)
      Else
         sqr_para_for_two_root = "F"
      End If
  Else
      sqr_para_for_two_root = "F"
  End If
  Exit Function
sqr_para_for_two_root_error:
sqr_para_for_two_root = "F"
End Function
Public Function solve_general_equation(ByVal s1$, ByVal S2$, unknown_element As String) As String
Dim ty(1) As Byte
Dim ts1(1) As String
Dim ts2(1) As String
ty(0) = string_type(s1, "", ts1(0), ts1(1), "")
ty(1) = string_type(S2, "", ts2(0), ts2(1), "")
If ty(0) = 3 And ty(1) = 3 Then
 solve_general_equation = solve_first_order_equation(minus_string( _
         time_string(ts1(0), ts2(1), False, False), _
           time_string(ts2(0), ts1(1), False, False), True, False), "0", unknown_element)
ElseIf ty(0) = 3 Then
    solve_general_equation = solve_first_order_equation(minus_string(ts1(0), _
           time_string(ts1(1), S2$, False, False), True, False), "0", unknown_element)
ElseIf ty(1) = 3 Then
 solve_general_equation = solve_first_order_equation(minus_string(time_string(s1$, _
           ts2(1), False, False), ts2(0), True, False), "0", unknown_element)
Else
 solve_general_equation = solve_first_order_equation(minus_string(s1$, S2$, True, False), _
    "0", unknown_element)
End If
End Function

Public Function simple_item_for_squre_root(ByVal temp_I As String) As String
Dim tn%, st%, la%
Dim ch1$
Dim ch2$
Dim ch3$
Dim ch4$
Dim t_i(2) As String
If Len(temp_I) = 0 Then
 simple_item_for_squre_root = "1"
ElseIf Len(temp_I) = 1 Then
 simple_item_for_squre_root = temp_I
Else
 st% = InStr(1, temp_I, "[", 0)
 If st% > 1 Then
  t_i(0) = Mid$(temp_I, 1, st% - 1)
 End If
 If st% > 0 Then
  t_i(1) = read_sqr_no_from_string(temp_I, st%, la%, "")
   st% = InStr(1, temp_I, "[", 0)
    If st% > la% Then
     t_i(0) = t_i(0) + Mid$(temp_I, 1, st% - 1)
    End If
   If st% > 0 Then
    t_i(2) = read_sqr_no_from_string(temp_I, st%, la%, "")
     If la% < Len(temp_I) Then
       t_i(0) = t_i(0) + Mid$(temp_I, la% + 1, Len(temp_I) - la%)
     End If
   Else
    t_i(2) = "1"
     t_i(0) = t_i(0) + Mid$(temp_I, la% + 1, Len(temp_I) - la%)
   End If
 Else
  t_i(0) = temp_I
  t_i(1) = "1"
  t_i(2) = "1"
 End If
  'ch1$ = Mid$(temp_I, Len(temp_I) - 1, 1)
   'ch2$ = Mid$(temp_I, Len(temp_I), 1)
 If t_i(1) <> "1" And t_i(2) <> "1" Then 'Asc(ch1$) < 0 And ch1$ <> "'" And ch1$ <> "\" And Asc(ch2$) < 0 And ch2$ <> "'" And ch2$ <> "\" Then
  If t_i(1) = t_i(2) Then
   If t_i(0) = "1" Then
    simple_item_for_squre_root = number_string(t_i(1))
   Else
    simple_item_for_squre_root = time_string( _
       simple_item_for_squre_root(t_i(0)), t_i(1), True, False)
   End If
    Exit Function
  Else
   ch1$ = time_string(t_i(1), t_i(2), True, False)
    Call set_squre_root_string(ch1$, ch1$)
       simple_item_for_squre_root = time_string( _
            simple_item_for_squre_root(ch3$), ch1$, True, False)
     Exit Function
  End If
Else 'If t_i(2) <> "1" Then 'Asc(ch2$) < 0 And ch2$ <> "'" And ch2$ <> "\" Then
' ch4$ = simple_item_for_squre_root(t_i(0))
 ' If InStr(2, ch4$, "+", 0) > 0 Or InStr(2, ch4$, "-", 0) > 0 Or _
       InStr(2, ch4$, "#", 0) > 0 Or InStr(2, ch4$, "@", 0) > 0 Then
  '    simple_item_for_squre_root = time_string(ch4$, t_i(2), True, False)
  'Else
   ' simple_item_for_squre_root = temp_I
  'End If
'ElseIf t_i(1) <> "1" Then 'Asc(ch1$) < 0 And ch1$ <> "'" And ch1$ <> "\" Then
 ' ch4$ = simple_item_for_squre_root(t_i(0))
  'If InStr(2, ch4$, "+", 0) > 0 Or InStr(2, ch4$, "-", 0) > 0 Or _
       InStr(2, ch4$, "#", 0) > 0 Or InStr(2, ch4$, "@", 0) > 0 Then
   '   simple_item_for_squre_root = time_string(time_string(ch4$, t_i(1), False, False), _
             ch2$, True, False)
  'Else
    simple_item_for_squre_root = temp_I
  'End If
'Else
simple_item_for_squre_root = temp_I
End If
End If
End Function

Public Function is_contain_x(ByVal s$, ByVal X$, ByVal st%) As Boolean
Dim i%
'Dim tn%
Dim ch As String
Dim sq_v As String
If InStr(st%, s$, X$, 0) > 0 Then
is_contain_x = True
Else
i% = InStr(st%, s$, "[", 0)
  If i% = 0 Then
   is_contain_x = False
  Else
      sq_v = read_sqr_no_from_string(s$, i%, st%, "")
        is_contain_x = is_contain_x(sq_v, X$, 1)
       If is_contain_x = True Then
         Exit Function
       Else
        If Len(s$) > st% Then
        is_contain_x = is_contain_x(s$, X$, st% + 1)
        Else
        is_contain_x = False
        End If
       End If
  End If
 End If
End Function
Public Function read_sqr_no_from_string(ByVal s As String, ByVal st%, la%, sqr$) As String
Dim no%
la% = InStr(st%, s, "]", 0)
If la% > st% Then
 sqr = Mid$(s, st%, la% - st% + 1)
 no% = val(Mid$(s, st% + 1, la% - st% - 1))
  If no% > 0 Then
  read_sqr_no_from_string = number_string(no%)
  End If
End If
End Function
Public Function read_sqr_from_item(ByVal ite As String, st%, s$, sqr_v$, sqr$) As Byte
Dim no%, i%, la% 'sqr$ [?] sqr_v
Dim ch$
st% = InStr(st%, ite, "[", 0)
If st% > 0 Then
 la% = InStr(st%, ite, "]", 0)
  If la% > st% Then
   sqr = Mid$(ite, st%, la% - st% + 1)
    no% = val(Mid$(ite, st% + 1, la% - st% - 1))
     If no% > 0 Then
     sqr_v = number_string(no%)
     Else
      sqr_v = "1"
     End If
     If st% > 1 Then
      s$ = Mid$(ite, 1, st% - 1)
     Else
      s$ = ""
     End If
     If la% < Len(ite) Then
      s$ = s$ + Mid$(ite, la% + 1, Len(ite) - la%)
     End If
     If s$ = "#1" Or s$ = "+1" Or s$ = "" Then
      s$ = "1"
     ElseIf s$ = "-" Or s$ = "@" Then
      s$ = "-1"
     End If
    End If
Else
 st% = InStr(1, ite, "'", 0)
 If st% > 0 Then
    s$ = Mid$(ite, 1, st% - 1)
     sqr_v$ = ""
      If Len(ite) > st% Then
       For i% = st% + 1 To Len(ite)
        ch$ = Mid$(ite, i%, 1)
         If ch$ <> "'" Then
          sqr_v$ = sqr_v$ + ch$
         Else
          s$ = s$ + Mid$(ite, i%, Len(ite) - i% + 1)
           GoTo read_sqr_from_item_next1
         End If
         If ch$ = ")" Then
          If i% < Len(ite) Then
           s$ = s$ + Mid$(ite, i% + 1, Len(ite) - i%)
          End If
           GoTo read_sqr_from_item_next1
         End If
       Next i%
      End If
read_sqr_from_item_next1:
 If sqr_v$ = "" Then
  sqr_v = "1"
 End If
 If s$ = "" Or s$ = "+" Or s$ = "#" Then
   s$ = "1"
 ElseIf s$ = "-" Or s$ = "@" Then
   s$ = "@1"
 End If
 sqr$ = "[0]"
 Else
 sqr_v$ = "1"
 s$ = ite
 sqr$ = "[0]"
 End If
End If
End Function

Public Function read_sqr_from_string(ByVal s As String, ty As Byte, s1 As String) As String
Dim i%, last_sq%
Dim ch As String
Dim ts_(3) As String
Dim ts$
s1$ = ""
For i% = 1 To Len(s)
  ch = Mid$(s, i%, 1)
   If ch = "[" Then
       ts$ = read_sqr_no_from_string(s, i%, i%, "")
        If InStr(2, ts$, "+", 0) > 0 Or InStr(2, ts$, "#", 0) > 0 Or _
            InStr(1, ts$, "-", 0) > 0 Or InStr(1, ts$, "@", 0) > 0 Then
         ts$ = "(" + ts$ + ")"
        End If
        last_sq% = last_sq% + 1
        If ty = 0 Then
         read_sqr_from_string = read_sqr_from_string + ts$
        Else
         read_sqr_from_string = read_sqr_from_string + _
            read_sqr_from_string(ts$, ty, s1)
        End If
          
   Else
        s1 = s1 + ch
   End If
Next i%
If s1 = "" Then
 s1 = "1"
End If
If read_sqr_from_string = "" Then
 read_sqr_from_string = "1"
End If
End Function

Public Function read_sqr_item_from_item(ByVal it$, ByVal i%, out_i%) As String
Dim j%
Dim ch$
 For j% = i% To Len(it$)
  ch$ = Mid$(it$, i%, 1)
   If ch$ = ")" Then
    read_sqr_item_from_item = read_sqr_item_from_item + ch$
     out_i% = j% + 1
      Exit Function
   ElseIf ch$ = "'" Then
     out_i% = j%
      Exit Function
   End If
  Next j%
End Function
Public Sub read_para_and_const_from_first_order_equation(ByVal eq$, ByVal unkown_element$, para$, con$)
Dim ts(2) As String
ts(0) = eq$
para$ = "0"
con$ = "0"
Do
Call string_type(ts(0), ts(1), "", "", ts(2))
 If InStr(1, ts(1), unkown_element$, 0) > 0 Then
  para$ = add_string(para$, divide_string(ts(1), unkown_element$, False, False), True, False)
 Else
  con$ = add_string(con$, ts(1), True, False)
 End If
 ts(0) = ts(2)
Loop Until ts(2) = ""
End Sub

Public Sub read_sqr_string_from_string(ByVal s$, ByVal sta%, en%, out_p$, out_item$, re$)
Dim i% 're$余
Dim ch$
Dim ty As Byte '=1 start
If sta% > 0 Then
re$ = Mid$(s$, 1, sta% - 1)
Else
re$ = ""
End If
out_p$ = ""
out_item$ = ""
For i% = sta% To Len(s$)
 ch$ = Mid$(s$, i%, 1)
  If ch$ = "'" Then
    ty = 1 '开始
  ElseIf ty = 1 Then
   If ch$ <> "|" Then
    If ch$ < "A" Then
     out_p$ = out_p$ & ch
    Else
     out_item$ = out_item$ & ch
    End If
   ElseIf ch$ = "|" Then
    re$ = re$ & Mid$(s$, i%, Len(s$) - i%)
     en% = i%
      GoTo read_sqr_string_from_string_mark0
   End If
  Else
    re$ = re$ & ch$
  End If
Next i%
  If en% = 0 Then
   en% = i% - 1
  End If
read_sqr_string_from_string_mark0:
  If re$ = "" Or re$ = "+" Or re$ = "#" Then
   re$ = "1"
  ElseIf re$ = "-" Or re$ = "@" Then
   re$ = "-1"
  End If
  If out_p$ = "" Then
     out_p$ = "1"
  End If
  If out_item$ = "" Then
     out_item$ = "1"
  End If
End Sub
Public Sub simple_two_long(int1 As Long, int2 As Long, int3 As Long)
Dim temp0, temp1, temp2, temp3 As Long
temp0 = int1
 If int1 < 0 Then    '　保证分母为正
  int1 = -int1
   int2 = -int2
 End If
 temp1 = Abs(int1)
  temp2 = Abs(int2)
   temp3 = temp2
 If temp1 <> 0 And temp2 <> 0 Then
   While temp3 <> 0
    temp2 = temp1
      temp1 = temp3
        temp3 = temp2 Mod temp1
    Wend
 int1 = int1 / temp1
  int2 = int2 / temp1
 End If
 int3 = temp0 / int1
End Sub
Public Function display_string_(ByVal s$, dis_ty As Byte) As String
'ty=1 原串 ,ty=0 全川
Dim i%, tn%, pi_n%
Dim ch As String
Dim ch1 As String
Dim ts(1) As String
Dim t As Byte
Dim tv As v_string
Dim has_brace As Boolean
If s$ = "" Then
   display_string_ = ""
    Exit Function
End If
If dis_ty = 0 Then
 display_string_ = "!" + s$ + "~"
Else
If InStr(1, s$, "!", 0) = 1 Then
 s$ = Mid$(s$, 2, Len(s$) - 1)
End If
If InStr(1, s$, "~", 0) = Len(s$) Then
 s$ = Mid$(s$, 1, Len(s$) - 1)
End If
has_brace = remove_brace(s$)
If InStr(1, s$, "X", 0) > 0 Then
     ts(0) = read_para_from_string_for_ietm(s$, "X", "")
     If ts(0) = "1" Then
     display_string_ = "[arrow\\u]x[arrow\\v]"
     ElseIf ts(0) = "-1" Then
     display_string_ = "-[arrow\\u]x[arrow\\v]"
     Else
      If InStr(2, ts(0), "+", 0) > 0 Or InStr(2, ts(0), "-", 0) > 0 Or _
            InStr(2, ts(0), "/", 0) > 0 Then
       display_string_ = "(" + ts(0) + ")" + "[arrow\\u]x[arrow\\v]"
      Else
       display_string_ = ts(0) + "[arrow\\u]x[arrow\\v]"
      End If
     End If
     If has_brace Then
      display_string_ = "(" + display_string_ + ")"
     End If
     Exit Function
ElseIf InStr(1, s$, "UU", 0) = 0 And InStr(1, s$, "UV", 0) = 0 And _
     InStr(1, s$, "VV", 0) = 0 Then
   If InStr(1, s$, "U", 0) > 0 Or InStr(1, s$, "V", 0) > 0 Then
     tv = from_string_to_v_string(s$)
     ts(0) = display_string_(tv.coord(0), dis_ty)
     ts(1) = display_string_(tv.coord(1), dis_ty)
     If ts(0) = "0" Or ts(0) = "" Then
        ts(0) = "0"
     ElseIf ts(0) = "1" Then
      ts(0) = "[arrow\\u]"
     ElseIf ts(0) = "-1" Then
      ts(0) = "-[arrow\\u]"
     ElseIf InStr(2, ts(0), "+", 0) = 0 And InStr(2, ts(0), "-", 0) = 0 And _
            InStr(2, ts(0), "/", 0) = 0 Then
      ts(0) = ts(0) + "[arrow\\u]"
     Else
      ts(0) = "(" + ts(0) + ")[arrow\\u]"
     End If
     If ts(1) = "0" Or ts(1) = "" Then
        ts(1) = "0"
     ElseIf ts(1) = "1" Then
      ts(1) = "[arrow\\v]"
     ElseIf ts(1) = "-1" Then
      ts(1) = "-[arrow\\v]"
     ElseIf InStr(2, ts(1), "+", 0) = 0 And InStr(2, ts(1), "-", 0) = 0 And _
            InStr(2, ts(1), "/", 0) = 0 Then
      ts(1) = ts(1) + "[arrow\\v]"
     Else
      ts(1) = "(" + ts(1) + ")[arrow\\v]"
     End If
     If ts(1) = "0" Then
        display_string_ = ts(0)
     ElseIf ts(0) = "0" Then
        display_string_ = ts(1)
     Else
        If Mid$(ts(1), 1, 1) = "-" Then
         display_string_ = ts(0) + ts(1)
        Else
         display_string_ = ts(0) + "+" + ts(1)
        End If
     End If
     If has_brace Then
        display_string_ = "(" + display_string_ + ")"
     End If
   Exit Function
   End If
End If

pi_n% = InStr(1, s$, "\", 0)
If pi_n% > 0 Then
 s$ = Mid$(s$, 1, Len(s$) - 1)
End If
t = string_type(s$, "", ts(0), ts(1), "")
If pi_n% > 0 And t = 3 Then
   display_string_ = display_string_(ts(0), 0) & LoadResString_(1455, "") & _
      "/" & display_string_(ts(1), 0)
Else
display_string_ = ""
For i% = 1 To Len(s$)
ch = Mid$(s$, i%, 1)
If ch = "~" Or ch = "!" Then
 display_string_ = display_string_
ElseIf ch = "'" Or ch = LoadResString_(1460, "") Then
 display_string_ = display_string_ + LoadResString_(1460, "")
ElseIf ch = "#" Then
 display_string_ = display_string_ + "+"
ElseIf ch = "@" Then
 display_string_ = display_string_ + "-"
ElseIf ch = "&" Then
 display_string_ = display_string_ + "/"
ElseIf ch = "\" Or ch = LoadResString_(1456, "") Then
 display_string_ = display_string_ + LoadResString_(1455, "")
ElseIf ch = "+" Then
 ch1 = Mid$(s$, i% + 1, 1)
 If ch1 = "@" Then
  display_string_ = display_string_ + "-"
   i% = i% + 1
 ElseIf ch1 = "#" Then
   display_string_ = display_string_ + "+"
   i% = i% + 1
 Else
  display_string_ = display_string_ + ch
 End If
ElseIf ch = "[" Then   '0 And ch <> "'" And ch <> "\" Then 'And ty = 1 Then
'tn% = from_char_to_no(ch)
 ch1 = read_sqr_no_from_string(s, i%, i%, "")
 'If InStr(2, ch1, "-", 0) > 0 Or InStr(2, ch1, "-", 0) > 0 Or InStr(2, ch1, "-", 0) > 0 Or _
      InStr(2, ch1, "-", 0) > 0 Then
 display_string_ = display_string_ + LoadResString_(1460, "") + "(" + _
   display_string_(ch1, 0) + ")"
Else
  display_string_ = display_string_ + ch
End If
Next i%
End If
End If
If has_brace Then
display_string_ = "(" + display_string_ + ")"
End If
End Function
Public Function simple_string(ByVal s As String) As String
Dim ts As String
ts = s
While InStr(1, ts, "*", 0) = 1 Or InStr(1, ts, "+", 0) = 1 Or _
  InStr(1, ts, "@", 0) = 1 Or InStr(1, ts, "#", 0) = 1 Or _
   InStr(1, ts, empty_char, 0) = 1
ts = Mid$(ts, 2, Len(ts) - 1)
Wend
While (InStr(1, ts, "@", 0) = 2 Or InStr(1, ts, "+", 0) = 2 Or _
 InStr(1, ts, "#", 0) = 2) And InStr(1, ts, "0", 0) = 1
  ts = Mid$(ts, 3, Len(ts) - 2)
Wend
While InStr(1, ts, "*", 0) = 2 And InStr(1, ts, "1", 0) = 1
  ts = Mid$(ts, 3, Len(ts) - 2)
Wend
While InStr(Len(ts), ts, "*", 0) = Len(ts) Or InStr(Len(ts), ts, "+", 0) = Len(ts) Or _
  InStr(Len(ts), ts, "@", 0) = Len(ts) Or InStr(Len(ts), ts, "#", 0) = Len(ts) Or _
   InStr(Len(ts), ts, empty_char, 0) = Len(ts)
ts = Mid$(ts, 1, Len(ts) - 1)
Wend
If Len(ts) > 1 Then
While (InStr(Len(ts) - 1, ts, "@", 0) = Len(ts) - 1 Or InStr(Len(ts) - 1, ts, "+", 0) = Len(ts) - 1 Or _
 InStr(Len(ts) - 1, ts, "#", 0) = Len(ts) - 1) And InStr(Len(ts), ts, "0", 0) = Len(ts)
  ts = Mid$(ts, 1, Len(ts) - 2)
Wend
While InStr(Len(ts) - 1, ts, "*", 0) = Len(ts) - 1 And InStr(Len(ts), ts, "1", 0) = Len(ts)
  ts = Mid$(1, 3, Len(ts) - 2)
Wend
End If
simple_string = ts
  End Function
Public Function read_para_from_equation_for_ietm(ByVal eq$, ByVal item$, re_eq$) As String
'从eq$读出item$的系数,re_eq$余项
Dim t_item$, t_item1$, t_eq$
Dim s(4) As String
 t_item$ = item$
 read_para_from_equation_for_ietm = "0"
 re_eq$ = "0"
If Len(t_item$) = 0 Then
 re_eq$ = eq$
ElseIf Len(t_item$) = 1 Then
s(0) = eq$
Do While (s(0) <> "")
 Call string_type(s(0), s(1), s(2), s(3), s(4))
  If InStr(1, s(3), t_item, 0) > 0 Then
   read_para_from_equation_for_ietm = add_string(read_para_from_equation_for_ietm, _
        time_string(s(2), divide_string(s(3), t_item$, False, False), False, False), True, False)
   Else
        re_eq$ = add_string(re_eq$, s(1), True, False)
  End If
  s(0) = s(4)
Loop
Else
 t_item1$ = Mid$(t_item$, 1, 1)
  t_item$ = Mid$(t_item$, 2, Len(t_item$) - 1)
   s(0) = read_para_from_equation_for_ietm(eq$, t_item1$, re_eq$)
    read_para_from_equation_for_ietm = _
      read_para_from_equation_for_ietm(s(0), t_item$, t_eq$)
       re_eq$ = add_string(re_eq$, time_string(t_eq$, t_item1$, False, False), _
             True, False)
End If
End Function

Public Sub peifang(ByVal s0$, ByVal s1$, S2$, re$)
'配方 s0$ 二次项目s1$一次项,s2$平方式,re$ 余项unknown$未知数
's0+s1=a$(s2$)^2+er$
Dim s(1) As String
Dim f  As String
If s0$ <> "1" Then
 s(0) = divide_string(s1$, s0$, True, False)
  Call peifang("1", s(0), S2$, s(1))
   re$ = time_string(s(1), s0$, True, False)
Else
S2 = divide_string(s1$, "2", True, False)
re$ = time_string(S2, S2, True, False)
re$ = time_string(re$, "-1", True, False)
End If
End Sub
Public Function find_first_item(s$, sig1 As String, sig2 As String) As Integer
Dim n1%, n2%, k%
Dim ch1$, ch2$
n1% = 1
n2% = InStr(2, s$, sig1, 0)
k% = InStr(2, s$, sig2, 0)
If k% > 0 And (n2% = 0 Or k% < n2%) Then
 n2% = k%
End If
'n2%第一的加减号
k% = InStr(1, s$, "(", 0)
If k% > 0 And (n2% = 0 Or k% < n2%) Then
  n1% = k%
   n2% = right_brace(s$, n1% + 1)
    If Len(s$) = n2% Then
     find_first_item = n2%
    Else
      ch1$ = Mid$(s$, n2% + 1, 1)
      If ch1$ = sig1 Or ch1$ = sig2 Then
       find_first_item = n2%
      Else
       ch2$ = Mid$(s$, n2% + 1, Len(s$) - n2%)
       find_first_item = n2% + find_first_item(ch2$, sig1, sig2)
      End If
    End If
ElseIf n2% > 0 Then
 find_first_item = n2% - 1
Else
 find_first_item = Len(s$)
End If
End Function
Public Function V_time_string(ByVal s1$, ByVal S2$) As String
Dim ty(1) As Byte
Dim s(9) As String
If InStr(1, s1$, "u", 0) = 0 And InStr(1, s1$, "v", 0) = 0 Then
 V_time_string = time_string(s1$, S2$, True, False)
ElseIf InStr(1, S2$, "u", 0) = 0 And InStr(1, S2$, "v", 0) = 0 Then
 V_time_string = time_string(s1$, S2$, True, False)
Else
 ty(0) = string_type(s1$, s(0), s(1), s(2), s(3))
 ty(1) = string_type(S2$, s(4), s(5), s(6), s(7))
 If ty(0) = 0 And ty(1) = 0 Then
    If s(3) = "" And s(7) = "" Then
      s(3) = time_para(s(1), s(5), True, False)
      s(7) = V_time_item(s(2), s(5))
      If s(7) = "0" Then
      V_time_string = "0"
      ElseIf s(3) = "1" Then
      V_time_string = s(7)
      Else
       If Mid$(s(7), 1, 1) = "-" Then
        V_time_string = "-" & s(3) & Mid$(s(7), 2, Len(s(7)) - 1)
       Else
        V_time_string = s(3) & s(7)
       End If
      End If
    ElseIf s(3) = "" Then
     V_time_string = add_string(V_time_string(s1$, s(4)), _
         V_time_string(s1$, s(7)), True, False)
    Else
     V_time_string = add_string(V_time_string(s(0), S2$), _
         V_time_string(s(3), S2$), True, False)
    End If
 ElseIf ty(0) = 0 And ty(1) = 3 Then
    s(3) = V_time_string(s1$, s(5))
    V_time_string = divide_string(s(3), s(6), True, False)
 ElseIf ty(0) = 3 And ty(1) = 0 Then
    V_time_string = time_string("-1", V_time_string(S2$, s1$), True, False)
 Else
    s(3) = V_time_string(s(1), s(5))
    s(7) = time_string(s(2), s(6), False, False)
    V_time_string = divide_string(s(3), s(7), True, False)

 End If
 
End If
End Function
Public Function V_time_item(ByVal i1$, ByVal i2$)
Dim l(1) As Integer
Dim ts(3) As String
l(0) = InStr(1, i1$, "v", 0)
If l(0) = 0 Then
l(0) = InStr(1, i1$, "u", 0)
If l(0) = 0 Then
 V_time_item = time_item(i1$, i2$)
End If
End If
l(1) = InStr(1, i2$, "v", 0)
If l(1) = 0 Then
l(1) = InStr(1, i2$, "u", 0)
If l(1) = 0 Then
 V_time_item = time_item(i1$, i2$)
End If
End If
ts(0) = Mid$(i1$, 1, l(0) - 1)
ts(1) = Mid$(i1$, l(0), Len(i1$) - l(0) + 1)
ts(2) = Mid$(i2$, 1, l(1) - 1)
ts(3) = Mid$(i2$, l(1), Len(i2$) - l(1) + 1)
If ts(0) = "" Then
    ts(0) = "1"
End If
If ts(2) = "" Then
    ts(2) = "1"
End If
 ts(0) = time_item(ts(0), ts(2))
If ts(1) = ts(3) Then
V_time_item = "0"
ElseIf ts(1) < ts(3) Then
 ts(1) = ts(1) & "**" & ts(3)
  If ts(0) = "1" Then
   V_time_item = ts(1)
  Else
   V_time_item = ts(0) & ts(1)
  End If
Else
 ts(1) = ts(3) & "**" & ts(1)
  If ts(0) = "1" Then
    V_time_item = "-" & ts(1)
  Else
   V_time_item = "-" & ts(0) & ts(1)
  End If
End If
End Function
Public Sub simple_two_int(int1 As Integer, int2 As Integer)
 Dim temp1 As Integer, temp2 As Integer, temp3 As Integer
 If int1 < 0 Then    '　保证分母为正
  int1 = -int1
   int2 = -int2
 End If
 temp1 = Abs(int1)
  temp2 = Abs(int2)
   temp3 = temp2
 If temp1 <> 0 And temp2 <> 0 Then
   While temp3 <> 0
    temp2 = temp1
      temp1 = temp3
        temp3 = temp2 Mod temp1
    Wend
 int1 = int1 / temp1
  int2 = int2 / temp1
 End If
End Sub
Public Function reduce_string_by_string(ts1$) As Boolean
'ts2 约化ts1
Dim i%
For i% = 1 To last_conditions.last_cond(1).relation_string_no
  If reduce_string_by_string0(ts1$, i%) Then
   Call add_conditions_to_record(relation_string_, i%, 0, 0, c_data_for_reduce)
  End If
Next i%
End Function
Public Function reduce_string_by_string0(ts1$, re_s%) As Boolean
Dim para$
 para$ = read_para_from_string_for_ietm(ts1$, relation_string(re_s%).data(0).element(0), "")
  If para$ <> "0" Then
     reduce_string_by_string0 = True
      ts1$ = minus_string(ts1$, time_string( _
         divide_string(relation_string(re_s%).data(0).relation_string, _
            relation_string(re_s%).data(0).para(0), False, False), para$, False, False), _
             True, False)
  End If
End Function
Public Function cross_time_v_item(s1$, S2$) As String
If s1$ = "U" And S2$ = "V" Then
   cross_time_v_item = "X"
ElseIf s1$ = "V" And S2$ = "U" Then
   cross_time_v_item = "-X"
ElseIf s1$ = "U" And S2$ = "U" Then
   cross_time_v_item = "0"
ElseIf s1$ = "V" And S2$ = "V" Then
   cross_time_v_item = "0"
Else
   cross_time_v_item = time_item(s1$, S2$)
End If
End Function
Public Function simple_v_string(ByVal v_s As String) As String '求出系数公因子
Dim pA(2) As String
If InStr(1, v_s, "UU", 0) > 0 Or InStr(1, v_s, "VU", 0) > 0 Or InStr(1, v_s, "VV", 0) > 0 Then
 pA(0) = read_para_from_string_for_ietm(v_s, "UU", "")
 pA(1) = read_para_from_string_for_ietm(v_s, "VU", "")
 pA(2) = read_para_from_string_for_ietm(v_s, "VV", "")
 Call simple_multi_string0(pA(0), pA(1), pA(2), "0", simple_v_string, False)
ElseIf InStr(1, v_s, "U", 0) > 0 Or InStr(1, v_s, "V", 0) > 0 Then
 pA(0) = read_para_from_string_for_ietm(v_s, "U", "")
 pA(1) = read_para_from_string_for_ietm(v_s, "V", "")
 Call simple_multi_string0(pA(0), pA(1), "0", "0", simple_v_string, False)
Else
 simple_v_string = v_s
End If
End Function
Public Function set_number_string(n_string As String, ByVal n%, Optional set_new_number As Boolean = False) As Integer
Dim i%
If n% = 0 Or n% > last_number_string Then
 If n% = 0 And set_new_number = False Then
  For i% = 1 To last_number_string
  If number_string(i%) = n_string Then
    set_number_string = i%
    Exit Function
  End If
  Next i%
  last_number_string = last_number_string + 1
 Else
  last_number_string = n%
 End If
ReDim Preserve number_string(last_number_string) As String
 set_number_string = last_number_string
Else
 set_number_string = n%
End If
 number_string(set_number_string) = n_string
End Function

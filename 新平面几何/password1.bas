Attribute VB_Name = "Module2"

Function cal_password_1(ByVal ser_no As String, computer_id As String) As String
Dim ch(2) As String
Dim out_string(2) As String
Dim i%, m%
Dim num(1) As Integer
Dim l As Long
If ser_no = "" Then
ser_no = "028-000001"
End If
'm% = InStr(1, ser_no, "-", 0)
'ch(0) = Mid$(ser_no, m% + 1, 6)
ch(0) = ser_no 'ch(0) + Mid$(ser_no, m% - 2, 2)
ch(1) = Trim(Mid$(computer_id, 1, 8))
ch(0) = set_protect_code_1(ch(0))
ch(1) = set_protect_code_1(ch(1))
ch(2) = ""
For i% = 1 To 8
ch(2) = ch(2) + Mid$(ch(0), i%, 1)
ch(2) = ch(2) + Mid$(ch(1), i%, 1)
Next i%
ch(2) = Mid$(ch(0), i%, 9) + ch(2) + Mid$(ch(0), i%, 10)
ch(1) = Mid$(ch(2), 1, 10)
ch(0) = Mid$(ch(2), 11, 18)
ch(0) = set_protect_code_1(ch(0))
ch(1) = set_protect_code_1(ch(1))
ch(2) = ""
For i% = 1 To 8
ch(2) = ch(2) + Mid$(ch(0), i%, 1)
ch(2) = ch(2) + Mid$(ch(1), i%, 1)
Next i%
ch(2) = Mid$(ch(1), i%, 9) + ch(2) + Mid$(ch(1), i%, 10)
ch(0) = Mid$(ch(2), 5, 14)
ch(1) = Mid$(ch(2), 1, 4) + Mid$(ch(2), 15, 18)
ch(0) = set_protect_code_1(ch(0))
ch(1) = set_protect_code_1(ch(1))
ch(2) = ""
For i% = 1 To 8
ch(2) = ch(2) + Mid$(ch(0), i%, 1)
ch(2) = ch(2) + Mid$(ch(1), i%, 1)
Next i%
ch(2) = Mid$(ch(0), i%, 9) + ch(2) + Mid$(ch(0), i%, 10)
ch(0) = Mid$(ch(2), 1, 8)
ch(1) = Mid$(ch(2), 9, 18)
ch(0) = set_protect_code_1(ch(0))
ch(1) = set_protect_code_1(ch(1))
ch(2) = ""
For i% = 1 To 8
ch(2) = ch(2) + Mid$(ch(0), 9 - i%, 1)
ch(2) = ch(2) + Mid$(ch(1), i%, 1)
Next i%
out_string(0) = Mid$(ch(2), 1, 4)
out_string(1) = Mid$(ch(2), 5, 4)
num(0) = val(out_string(0))
num(1) = val(out_string(1))
out_string(0) = Trim(str(num(0) + num(1)))
If Len(out_string(0)) = 1 Then
out_string(0) = "000" + out_string(0)
ElseIf Len(out_string(0)) = 2 Then
out_string(0) = "00" + out_string(0)
ElseIf Len(out_string(0)) = 3 Then
out_string(0) = "0" + out_string(0)
ElseIf Len(out_string(0)) > 4 Then
out_string(0) = Mid$(out_string(0), 1, 4)
End If
out_string(2) = Mid$(ch(2), 9, 8)
cal_password_1 = out_string(0) + out_string(2)
End Function

Public Function cal_protect_code_1(in_ch As String, time%) As String
Dim l!, l_!
Dim i%
  l! = Asc(Mid$(in_ch, 1, 1)) Mod 100 '四位数
  l! = l! * 100 + (Asc(Mid$(in_ch, 2, 1)) Mod 100)
l! = function_for_protect_1(l!) '迭代
For i% = 1 To time% - 1 'time% 次迭代
l! = Abs(l!) Mod 10000
l! = function_for_protect_1(l!)
Next i%
l! = Abs(l!)
l_! = l! Mod 100
l! = (l_! + (l! - l_!) / 100) Mod 100
cal_protect_code_1 = Trim(str(l!))
If Len(cal_protect_code_1) = 1 Then
  cal_protect_code_1 = "0" + cal_protect_code_1
End If
End Function

Public Function function_for_protect_1(i!) As Long
function_for_protect_1 = -4 * i! * i! + 14 * i! + 6
End Function
Public Function set_protect_code_1(computer_id As String) As String
Dim i%
Dim temp_string(4) As String * 2
'两位分组
temp_string(0) = Mid$(computer_id, 1, 2)
temp_string(1) = Mid$(computer_id, 3, 2)
temp_string(2) = Mid$(computer_id, 5, 2)
temp_string(3) = Mid$(computer_id, 7, 2)
If Len(computer_id) > 8 Then
temp_string(4) = Mid$(computer_id, 9, 2)
End If
set_protect_code_1 = cal_protect_code_1(temp_string(0), 5)
If temp_string(1) = temp_string(0) Then
 set_protect_code_1 = set_protect_code_1 + cal_protect_code_1(temp_string(1), 6)
  If temp_string(2) = temp_string(1) Then '2=1=0
   set_protect_code_1 = cal_protect_code_1(temp_string(2), 7) + set_protect_code_1
    If temp_string(3) = temp_string(2) Then '3=2=1=0
     set_protect_code_1 = set_protect_code_1 + cal_protect_code_1(temp_string(3), 8)
    Else
     set_protect_code_1 = set_protect_code_1 + cal_protect_code_1(temp_string(3), 5)
    End If
  Else '0=1<>2
   set_protect_code_1 = cal_protect_code_1(temp_string(2), 5) + set_protect_code_1
    If temp_string(3) = temp_string(2) Or temp_string(3) = temp_string(1) Then '0=1 <>2,3=2or 3=0
     set_protect_code_1 = set_protect_code_1 + cal_protect_code_1(temp_string(3), 7)
    Else
     set_protect_code_1 = set_protect_code_1 + cal_protect_code_1(temp_string(3), 5)
    End If
  End If
Else
 set_protect_code_1 = set_protect_code_1 + cal_protect_code_1(temp_string(1), 5)
  If temp_string(2) = temp_string(1) Or temp_string(2) = temp_string(0) Then '
   set_protect_code_1 = cal_protect_code_1(temp_string(2), 6) + set_protect_code_1
   If temp_string(3) = temp_string(2) Or temp_string(3) = temp_string(1) Or _
         temp_string(3) = temp_string(0) Then
    set_protect_code_1 = set_protect_code_1 + cal_protect_code_1(temp_string(3), 7)
   Else
    set_protect_code_1 = set_protect_code_1 + cal_protect_code_1(temp_string(3), 6)
   End If
  Else
   set_protect_code_1 = cal_protect_code_1(temp_string(2), 5) + set_protect_code_1
    If temp_string(3) = temp_string(2) Or temp_string(3) = temp_string(1) Or _
     temp_string(3) = temp_string(0) Then
    set_protect_code_1 = set_protect_code_1 + cal_protect_code_1(temp_string(3), 6)
    Else
     set_protect_code_1 = set_protect_code_1 + cal_protect_code_1(temp_string(3), 5)
    End If
  End If
End If
If temp_string(4) <> "" Then
 set_protect_code_1 = set_protect_code_1 + cal_protect_code_1(temp_string(4), 5)
End If
End Function




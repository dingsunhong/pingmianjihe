Attribute VB_Name = "Module1"
Option Explicit

Public Function set_protect_code_for_cp(computer_id As String) As String
Dim i%
Dim temp_string(3) As String * 2
temp_string(0) = Mid$(computer_id, 1, 2)
temp_string(1) = Mid$(computer_id, 3, 2)
temp_string(2) = Mid$(computer_id, 5, 2)
temp_string(3) = Mid$(computer_id, 7, 2)
set_protect_code_for_cp = cal_protect_code_for_cp(temp_string(0), 5)
If temp_string(1) = temp_string(0) Then
 set_protect_code_for_cp = set_protect_code_for_cp + cal_protect_code_for_cp(temp_string(1), 6)
  If temp_string(2) = temp_string(1) Then '2=1=0
   set_protect_code_for_cp = cal_protect_code_for_cp(temp_string(2), 7) + set_protect_code_for_cp
    If temp_string(3) = temp_string(2) Then '3=2=1=0
     set_protect_code_for_cp = set_protect_code_for_cp + cal_protect_code_for_cp(temp_string(3), 8)
    Else
     set_protect_code_for_cp = set_protect_code_for_cp + cal_protect_code_for_cp(temp_string(3), 5)
    End If
  Else '0=1<>2
   set_protect_code_for_cp = cal_protect_code_for_cp(temp_string(2), 5) + set_protect_code_for_cp
    If temp_string(3) = temp_string(2) Or temp_string(3) = temp_string(1) Then '0=1 <>2,3=2or 3=0
     set_protect_code_for_cp = set_protect_code_for_cp + cal_protect_code_for_cp(temp_string(3), 7)
    Else
     set_protect_code_for_cp = set_protect_code_for_cp + cal_protect_code_for_cp(temp_string(3), 5)
    End If
  End If
Else
 set_protect_code_for_cp = set_protect_code_for_cp + cal_protect_code_for_cp(temp_string(1), 5)
  If temp_string(2) = temp_string(1) Or temp_string(2) = temp_string(0) Then '
   set_protect_code_for_cp = cal_protect_code_for_cp(temp_string(2), 6) + set_protect_code_for_cp
   If temp_string(3) = temp_string(2) Or temp_string(3) = temp_string(1) Or _
         temp_string(3) = temp_string(0) Then
    set_protect_code_for_cp = set_protect_code_for_cp + cal_protect_code_for_cp(temp_string(3), 7)
   Else
    set_protect_code_for_cp = set_protect_code_for_cp + cal_protect_code_for_cp(temp_string(3), 6)
   End If
  Else
   set_protect_code_for_cp = cal_protect_code_for_cp(temp_string(2), 5) + set_protect_code_for_cp
    If temp_string(3) = temp_string(2) Or temp_string(3) = temp_string(1) Or _
     temp_string(3) = temp_string(0) Then
    set_protect_code_for_cp = set_protect_code_for_cp + cal_protect_code_for_cp(temp_string(3), 6)
    Else
     set_protect_code_for_cp = set_protect_code_for_cp + cal_protect_code_for_cp(temp_string(3), 5)
    End If
  End If
End If
End Function

Public Function cal_protect_code_for_cp(in_ch As String, time%) As String
Dim l!
Dim i%
  l! = l! + Asc(Mid$(in_ch, 1, 1)) Mod 100
  l! = l! * 100 + (Asc(Mid$(in_ch, 2, 1)) Mod 100)
l! = function_for_protect_for_cp(l!)
For i% = 1 To time% - 1
l! = Abs(l!) Mod 10001
l! = function_for_protect_for_cp(l!)
Next i%
l! = Abs(l!) Mod 100
cal_protect_code_for_cp = Trim(str(l!))
If Len(cal_protect_code_for_cp) = 1 Then
  cal_protect_code_for_cp = "0" + cal_protect_code_for_cp
End If
End Function

Public Function function_for_protect_for_cp(i!) As Long
function_for_protect_for_cp = -3 * i! * i! + 15 * i! + 5
End Function





Attribute VB_Name = "input_for_wenti"
Option Explicit
Public Function get_last_used_char(ByVal w%, ByVal n%) As Integer
Dim i%, j%
For i% = 0 To w% - 1
 For j% = 0 To 10
    If C_display_wenti.m_condition(i%, j%) >= "A" And _
           C_display_wenti.m_condition(i%, j%) <= "Z" Then
      get_last_used_char = max(get_last_used_char, _
          char_number(C_display_wenti.m_condition(i%, j%)))
    End If
 Next j%
Next i%
For i% = 0 To n%
 If C_display_wenti.m_condition(w%, i%) >= "A" And _
     C_display_wenti.m_condition(w%, i%) <= "Z" Then
     get_last_used_char = max(get_last_used_char, _
          char_number(C_display_wenti.m_condition(w%, i%)))
 End If
Next i%
End Function
Public Function chose_wenti_no_(ByVal Y%, w_n%)
w_n% = Int(Y% / 16) - 1
End Function
Public Function char_number(ch$) As Integer
Dim i%
For i% = 1 To last_conditions.last_cond(1).point_no
 If m_poi(i%).data(0).data0.name = ch$ Then
  char_number = i%
   Exit Function
 End If
Next i%
End Function
Public Sub input_key_down(KeyCode As Integer)
End Sub
Public Sub input_key_press(KeyAscii As Integer)
End Sub

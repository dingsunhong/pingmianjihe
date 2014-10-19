Attribute VB_Name = "set_test_vision"
'Private Declare Function WHDFingerPrint Lib "try.dll" () As HDFingerPrint
'PrPrivate Declare Function WHDFingerPrint Lib "temp_inform.dll" () As HDFingerPrint
'Private Declare Function my_sysinformation Lib "temp_inform.dll" () As SysInform_records
'Private Declare Function my_sysinformation Lib "My_SysInforLib.dll" () As SysInform_records

Option Explicit
Global c_line1_5_enabled As Boolean
Global c_choose_enabled As Boolean
Global print_enabled As Boolean
Global Const file_name1 = "\media\The Microsoft Sound.wav"
Global Const file_name2 = "\System32\Drivers\Update.sys"

Public Function set_protect_code(computer_id As String) As String
Dim i%
Dim temp_string(3) As String * 2
temp_string(0) = Mid$(computer_id, 1, 2)
temp_string(1) = Mid$(computer_id, 3, 2)
temp_string(2) = Mid$(computer_id, 5, 2)
temp_string(3) = Mid$(computer_id, 7, 2)
set_protect_code = cal_protect_code(temp_string(0), 5)
If temp_string(1) = temp_string(0) Then
 set_protect_code = set_protect_code + cal_protect_code(temp_string(1), 6)
  If temp_string(2) = temp_string(1) Then '2=1=0
   set_protect_code = cal_protect_code(temp_string(2), 7) + set_protect_code
    If temp_string(3) = temp_string(2) Then '3=2=1=0
     set_protect_code = set_protect_code + cal_protect_code(temp_string(3), 8)
    Else
     set_protect_code = set_protect_code + cal_protect_code(temp_string(3), 5)
    End If
  Else '0=1<>2
   set_protect_code = cal_protect_code(temp_string(2), 5) + set_protect_code
    If temp_string(3) = temp_string(2) Or temp_string(3) = temp_string(1) Then '0=1 <>2,3=2or 3=0
     set_protect_code = set_protect_code + cal_protect_code(temp_string(3), 7)
    Else
     set_protect_code = set_protect_code + cal_protect_code(temp_string(3), 5)
    End If
  End If
Else
 set_protect_code = set_protect_code + cal_protect_code(temp_string(1), 5)
  If temp_string(2) = temp_string(1) Or temp_string(2) = temp_string(0) Then '
   set_protect_code = cal_protect_code(temp_string(2), 6) + set_protect_code
   If temp_string(3) = temp_string(2) Or temp_string(3) = temp_string(1) Or _
         temp_string(3) = temp_string(0) Then
    set_protect_code = set_protect_code + cal_protect_code(temp_string(3), 7)
   Else
    set_protect_code = set_protect_code + cal_protect_code(temp_string(3), 6)
   End If
  Else
   set_protect_code = cal_protect_code(temp_string(2), 5) + set_protect_code
    If temp_string(3) = temp_string(2) Or temp_string(3) = temp_string(1) Or _
     temp_string(3) = temp_string(0) Then
    set_protect_code = set_protect_code + cal_protect_code(temp_string(3), 6)
    Else
     set_protect_code = set_protect_code + cal_protect_code(temp_string(3), 5)
    End If
  End If
End If
End Function



Sub set_enabled(ByVal ty As Boolean)
'ty=true 正式版
 c_line1_5_enabled = ty
 c_choose_enabled = ty
 print_enabled = ty
  '*******
If protect_data.install_statue = "F" Then '使用期满
'welcome.ZOrder 0
'welcome.Show 1
ElseIf protect_data.install_statue = "T" Then '正式版,中途更改系统
'If is_temp_inform_dll_open = False Then
' If protect_data.pass_word <> _
 '        cal_password(protect_data.serial_no, computer_id) Then
 ' protect_data.install_statue = "F"
 '  ty = False
' End If
'End If
End If
 MDIForm1.auto.Enabled = ty
 MDIForm1.chose_law.Enabled = ty
 MDIForm1.c_angle2.Enabled = ty
 MDIForm1.c_angle3.Enabled = ty
 MDIForm1.c_angle4.Enabled = ty
 MDIForm1.c_choose.Enabled = ty
 MDIForm1.c_circle.Enabled = ty
 MDIForm1.c_cal4.Enabled = ty
 MDIForm1.c_cal5.Enabled = ty
 MDIForm1.c_cal6.Enabled = ty
 MDIForm1.c_cal7.Enabled = ty
 MDIForm1.c_cal8.Enabled = ty
 MDIForm1.c_cal9.Enabled = ty
 MDIForm1.c_calA.Enabled = ty
 MDIForm1.c_calB.Enabled = ty
 MDIForm1.c_calC.Enabled = ty
 MDIForm1.c_line1_5.Enabled = ty
 MDIForm1.c_line2.Enabled = ty
 MDIForm1.c_line5.Enabled = ty
 MDIForm1.c_multi.Enabled = ty
 MDIForm1.c_triangle3.Enabled = ty
 MDIForm1.c_triangle4.Enabled = ty
 MDIForm1.c_triangle5.Enabled = ty
 MDIForm1.c_triangle6.Enabled = ty
 MDIForm1.c_triangle7.Enabled = ty
 MDIForm1.d_circle.Enabled = ty
 MDIForm1.d_circle_without_center.Enabled = ty
 MDIForm1.e_polygon3.Enabled = ty
 MDIForm1.e_polygon4.Enabled = ty
 MDIForm1.e_polygon5.Enabled = ty
 MDIForm1.e_polygon6.Enabled = ty
 MDIForm1.move_picture.Enabled = ty
 MDIForm1.re_name.Enabled = ty
 MDIForm1.set_opera_type = ty
 Print_Form.print1.Enabled = ty
 MDIForm1.tangent_circle.Enabled = ty
 MDIForm1.tangent_line.Enabled = ty
 MDIForm1.set_picture_for_change.Enabled = ty
 'MDIForm1.set_polygon_for_change.Enabled = ty
 'MDIForm1.set_circle_for_change.Enabled = ty
 MDIForm1.re_name_all.Enabled = ty
 MDIForm1.re_name_one.Enabled = ty
 MDIForm1.remove_point_.Enabled = ty
 MDIForm1.S5.Enabled = ty
 MDIForm1.S4.Enabled = ty
 MDIForm1.S61.Enabled = ty
 MDIForm1.S62.Enabled = ty
 MDIForm1.c_line1_5 = ty
 MDIForm1.c_choose = ty
 MDIForm1.S6E.Enabled = ty
 MDIForm1.S6F.Enabled = ty
 MDIForm1.S6G.Enabled = ty
 MDIForm1.S6H.Enabled = ty
 MDIForm1.S63.Enabled = ty
 MDIForm1.S64.Enabled = ty
 MDIForm1.S65.Enabled = ty
 MDIForm1.S66.Enabled = ty
 MDIForm1.S67.Enabled = ty
 MDIForm1.S68.Enabled = ty
 MDIForm1.S69.Enabled = ty
 MDIForm1.S6A.Enabled = ty
 MDIForm1.S6B.Enabled = ty
 MDIForm1.S6C.Enabled = ty
 MDIForm1.S6D.Enabled = ty
 MDIForm1.s1G.Enabled = ty
 MDIForm1.s1C.Enabled = ty
 MDIForm1.s1D.Enabled = ty
 MDIForm1.s1A.Enabled = ty
 'MDIForm1.s13.Enabled = ty
 MDIForm1.s1B.Enabled = ty
 MDIForm1.s18.Enabled = ty
 MDIForm1.s1D.Enabled = ty
 MDIForm1.S24.Enabled = ty
 MDIForm1.S32.Enabled = ty
 MDIForm1.S34.Enabled = ty
 MDIForm1.Open.Enabled = ty
 MDIForm1.save.Enabled = ty
 MDIForm1.savee.Enabled = ty
 MDIForm1.save_as.Enabled = ty
 MDIForm1.ratio_point.Enabled = ty
 MDIForm1.line_given_length.Enabled = ty
 MDIForm1.equal_line.Enabled = ty
 MDIForm1.equal_angle_line.Enabled = ty
 MDIForm1.draw_circle.Enabled = ty
 MDIForm1.verti_mid_line.Enabled = ty
 MDIForm1.E_polygon.Enabled = ty
 
End Sub
Public Function set_protect_data_to_string(pro_date As protect_data_type) As String
Dim install_date_string As String * 8
Dim used_time_string As String
Dim install_statue_string As String
'Dim install_date_string As String
Dim input_pass_word_time As String
Dim l%
set_protect_data_to_string = ""
used_time_string = str(pro_date.used_time)
used_time_string = Trim(used_time_string)
If Len(used_time_string) = 0 Then
 used_time_string = "0000"
ElseIf Len(used_time_string) = 1 Then
 used_time_string = "000" + used_time_string
ElseIf Len(used_time_string) = 2 Then
 used_time_string = "00" + used_time_string
ElseIf Len(used_time_string) = 3 Then
 used_time_string = "0" + used_time_string
End If
install_date_string = read_date_to_string(pro_date.install_date)
input_pass_word_time = Trim(str(pro_date.input_pass_word_time))
l% = Len(input_pass_word_time)
If l% = 0 Then
input_pass_word_time = "0000"
ElseIf l% = 1 Then
input_pass_word_time = "000" + input_pass_word_time
ElseIf l% = 2 Then
input_pass_word_time = "00" + input_pass_word_time
ElseIf l% = 3 Then
input_pass_word_time = "0" + input_pass_word_time
End If
If Len(pro_date.serial_no) > 10 Then
 pro_date.serial_no = Mid$(pro_date.serial_no, 2, 10)
End If
set_protect_data_to_string = pro_date.id + _
           protect_data_change(pro_date.computer_id) + protect_data_change(pro_date.pass_word) + _
          input_pass_word_time + install_date_string + used_time_string + pro_date.install_statue + _
           protect_data_change(pro_date.serial_no) + protect_data_change(pro_date.pass_word_for_teacher)
End Function

Public Function regist(ty As Byte) As Boolean
inform.Hide
If ty = 0 Then
frmLogin.Label1 = LoadResString_(1835, "")
Else
frmLogin.Label1 = LoadResString_(1840, "")
End If
frmLogin.ZOrder 0
frmLogin.LabelUserName.Caption = computer_id_
frmLogin.Text3.text = protect_data.serial_no
'frmLogin.txtPassword.SetFocus
frmLogin.Show 1
regist = LoginSucceeded
MDIForm1.Timer2.Enabled = True
End Function


Public Function read_date_to_string(ByVal right_date As String) As String
Dim ch As String
Dim m%
    If Mid$(right_date, 1, 1) = "#" Then
    right_date = Mid$(right_date, 2, Len(right_date) - 1)
    End If
    m% = InStr(1, right_date, "-", 0)
    If m% = 0 Then
     m% = InStr(1, right_date, "/", 0)
    End If
     ch = Mid$(right_date, 1, m%)
      If Len(ch) = 2 Then
       read_date_to_string = "0" + ch
      ElseIf Len(ch) = 3 Then
       read_date_to_string = ch
      Else
       read_date_to_string = Mid$(ch, Len(ch) - 2, 3)
      End If
       right_date = Mid$(right_date, m% + 1, Len(right_date) - m%)
     m% = InStr(1, right_date, "-", 0)
     If m% = 0 Then
      m% = InStr(1, right_date, "/", 0)
     End If
      ch = Mid$(right_date, 1, m%)
      If Len(ch) = 2 Then
       read_date_to_string = read_date_to_string + "0" + ch
      Else
       read_date_to_string = read_date_to_string + ch
      End If
      ch = Trim(Mid$(right_date, m% + 1, 2))
      If Len(ch) = 1 Then
       read_date_to_string = read_date_to_string + "0" + ch
      Else
       read_date_to_string = read_date_to_string + ch
      End If

End Function

Private Function read_data_from_protect_string(pro_string As String) As protect_data_type
Dim install_date_string As String * 8
Dim used_time_string As String * 4
Dim input_pass_word_time As String * 4
Dim records As Tsysinfor
read_data_from_protect_string.id = Mid$(pro_string, 1, 8)
If read_data_from_protect_string.id = "31304619" Then
'On Error GoTo set_date
read_data_from_protect_string.computer_id = protect_data_change(Mid$(pro_string, 9, 8))
read_data_from_protect_string.pass_word = protect_data_change(Mid$(pro_string, 17, 12))
input_pass_word_time = Mid$(pro_string, 29, 4)
read_data_from_protect_string.input_pass_word_time = val(input_pass_word_time)
read_data_from_protect_string.install_date = Mid$(pro_string, 33, 8)
used_time_string = Mid$(pro_string, 41, 4)
read_data_from_protect_string.used_time = val(used_time_string)
read_data_from_protect_string.install_statue = Mid$(pro_string, 45, 1)
read_data_from_protect_string.serial_no = protect_data_change(Mid$(pro_string, 46, 10))
read_data_from_protect_string.pass_word_for_teacher = protect_data_change(Mid$(pro_string, 56, 5))
'If Mid$(read_data_from_protect_string.serial_no, 1, 1) <> "0" Then
'    read_data_from_protect_string.serial_no = _
'      "0" + read_data_from_protect_string.serial_no
'End If
Exit Function
set_date:
read_data_from_protect_string.install_date = Now
'read_data_from_protect_string.computer_id = computer_id
read_data_from_protect_string.pass_word = set_protect_code(read_data_from_protect_string.pass_word)
read_data_from_protect_string.input_pass_word_time = 0
read_data_from_protect_string.install_statue = "S"
read_data_from_protect_string.used_time = 0
read_data_from_protect_string.serial_no = Mid$(pro_string, 46, 10)
read_data_from_protect_string.pass_word_for_teacher = Mid$(pro_string, 56, 5)
'If Mid$(read_data_from_protect_string.serial_no, 1, 1) <> "0" Then
'    read_data_from_protect_string.serial_no = _
      "0" + read_data_from_protect_string.serial_no
'End If
End If
End Function

Public Sub put_filetime(n%)
Dim sys_T As SYS_TIME
Dim File_T As FileTime
If n% = 0 Then
      protect_sys_time(0).Millsecond = protect_data.input_pass_word_time
      Call SystemTimeToFileTime(protect_sys_time(0), File_T)
      FileHandle(0) = OpenFile(protect_file(0), lpReOpenBuff, 2)
       Call SetFileTime(FileHandle(0), File_T, _
                     protect_file_time(0).ftLastAccessTime, _
                       protect_file_time(0).ftLastWriteTime)
       Call lclose(FileHandle(0))
ElseIf n% = 1 Then
sys_T = protect_sys_time(1)
sys_T.Hour = protect_sys_time(0).Hour
sys_T.Minite = protect_sys_time(0).Minite
sys_T.Second = protect_sys_time(0).Second
sys_T.Millsecond = protect_sys_time(0).Millsecond
      FileHandle(1) = OpenFile(protect_file(1), lpReOpenBuff, 2)
      Call SystemTimeToFileTime(sys_T, File_T)
       Call SetFileTime(FileHandle(1), File_T, _
                    protect_file_time(1).ftLastAccessTime, _
                      protect_file_time(1).ftLastWriteTime)
      Call lclose(FileHandle(1))
ElseIf n% = 2 Then
'protect_sys_time(3).Millsecond = protect_data.input_pass_word_time
protect_sys_time(2).Millsecond = protect_data.used_time
Call SystemTimeToFileTime(protect_sys_time(2), protect_file_time(2).ftCreationTime)
Call SystemTimeToFileTime(protect_sys_time(3), protect_file_time(2).ftLastWriteTime)
      FileHandle(0) = OpenFile(protect_file(2), lpReOpenBuff, 2)
       Call SetFileTime(FileHandle(0), protect_file_time(2).ftCreationTime, _
                     protect_file_time(2).ftLastAccessTime, _
                       protect_file_time(2).ftLastWriteTime)
       Call lclose(FileHandle(0))
End If
End Sub

Public Sub read_protect_file_time()

End Sub

Public Function DateDiff_(data_s1$, data_s2$) As Integer
Dim D1(2) As Integer
Dim D2(2) As Integer
D1(0) = val(Mid$(data_s1$, 1, 2))
D1(1) = val(Mid$(data_s1$, 4, 2))
D1(2) = val(Mid$(data_s1$, 7, 2))
D2(0) = val(Mid$(data_s2$, 1, 2))
D2(1) = val(Mid$(data_s2$, 4, 2))
D2(2) = val(Mid$(data_s2$, 7, 2))
If D1(0) - D2(0) = 0 Then '同年
 If D1(1) - D2(1) = 0 Then '同月
  DateDiff_ = D1(0) - D2(0)
   If DateDiff_ < 0 Then
    DateDiff_ = 20
   End If
 ElseIf D1(1) - D2(1) = 1 Then '不同月
   DateDiff_ = D1(2) + 30 - D2(2)
   If DateDiff_ < 0 Then
    DateDiff_ = 20
   End If
 Else
    DateDiff_ = 20
 DateDiff_ = 20
 End If
ElseIf D1(0) - D2(0) = 1 Then '不同年
 If D1(1) + 12 - D2(1) = 1 Then
   DateDiff_ = D1(2) + 30 - D2(2)
   If DateDiff_ < 0 Then
    DateDiff_ = 20
   End If
 Else
    DateDiff_ = 20
 End If
Else
DateDiff_ = 20
End If

End Function

Public Sub calculate_run_time() '
Dim data As String * 55
Dim dir$
Dim id As String
'Dim F_time(2) As FileTime
'*************
'On  Error goto calculate_run_time_mark0
 If protect_data.id = "31304619" Then
   If protect_data.install_statue <> "T" Then
    If protect_data.used_time < 300 And _
         protect_data.install_statue = "S" Then
        protect_data.used_time = protect_data.used_time + 1
       If protect_data.used_time >= 300 Then
        protect_data.install_statue = "F"
       End If
    ElseIf DateDiff_(read_date_to_string(DateValue(Date)), protect_data.install_date) > 15 Or _
       protect_data.used_time > 300 Then
        '未认证,超过认证期,成为试用版
          protect_data.install_statue = "F"
           '设置使用版
           End
                   'Call set_enabled(False)
    End If
   End If
       Open protect_file(0) For Binary As #5
     data = set_protect_data_to_string(protect_data)
      Put #5, 103, data
       Close #5
      Call put_filetime(0)
       Open protect_file(1) For Binary As #6
               Put #6, 1, data
       Close #6
      Call put_filetime(1)
      If protect_data.install_statue = "F" Then
         protect_sys_time(2).Millsecond = 300
      ElseIf protect_data.install_statue = "S" Then
         protect_sys_time(2).Millsecond = protect_data.used_time
      ElseIf protect_data.install_statue = "T" Then
         protect_sys_time(2).Millsecond = 301
      End If
      Call put_filetime(2)
Else
     'Unload MDIForm1 '不是嘉科软件
End If
If protect_data.install_statue = "F" Then
Call set_enabled(False)
End If
calculate_run_time_mark0:
End Sub
Public Function protect_data_change0(ByVal s As String) As String
If s = "0" Then
protect_data_change0 = "3"
ElseIf s = "1" Then
protect_data_change0 = "7"
ElseIf s = "2" Then
protect_data_change0 = "5"
ElseIf s = "3" Then
protect_data_change0 = "0"
ElseIf s = "4" Then
protect_data_change0 = "6"
ElseIf s = "5" Then
protect_data_change0 = "2"
ElseIf s = "6" Then
protect_data_change0 = "4"
ElseIf s = "7" Then
protect_data_change0 = "1"
ElseIf s = "8" Then
protect_data_change0 = "9"
ElseIf s = "9" Then
protect_data_change0 = "8"
ElseIf Asc(s) >= 64 And Asc(s) <= 115 Then
protect_data_change0 = Chr(179 - Asc(s))
 Else
protect_data_change0 = s
End If
End Function
Public Function protect_data_change(ByVal s As String) As String
Dim i%
For i% = 1 To Len(s)
protect_data_change = protect_data_change + protect_data_change0(Mid$(s, i%, 1))
Next i%
End Function
Function cal_password(ByVal ser_no As String, computer_id As String) As String
Dim ch(2) As String
Dim out_string(2) As String
Dim i%, m%
Dim num(1) As Integer
Dim l As Long
If ser_no = "" Then
ser_no = "028-000001"
End If
'm% = InStr(1, ser_no, "-", 0)
ch(0) = Mid$(ser_no, m% + 1, 6)
ch(0) = ser_no 'ch(0) + Mid$(ser_no, m% - 2, 2)
ch(1) = Trim(Mid$(computer_id, 1, 8))
ch(0) = set_protect_code(ch(0))
ch(1) = set_protect_code(ch(1))
ch(2) = ""
For i% = 1 To 8
ch(2) = ch(2) + Mid$(ch(0), i%, 1)
ch(2) = ch(2) + Mid$(ch(1), 9 - i%, 1)
Next i%
ch(0) = Mid$(ch(2), 1, 8)
ch(1) = Mid$(ch(2), 7, 8)
ch(0) = set_protect_code(ch(0))
ch(1) = set_protect_code(ch(1))
ch(2) = ""
For i% = 1 To 8
ch(2) = ch(2) + Mid$(ch(0), 9 - i%, 1)
ch(2) = ch(2) + Mid$(ch(1), i%, 1)
Next i%
ch(0) = Mid$(ch(2), 1, 8)
ch(1) = Mid$(ch(2), 7, 8)
ch(0) = set_protect_code(ch(0))
ch(1) = set_protect_code(ch(1))
ch(2) = ""
For i% = 1 To 8
ch(2) = ch(2) + Mid$(ch(0), i%, 1)
ch(2) = ch(2) + Mid$(ch(1), 9 - i%, 1)
Next i%
ch(0) = Mid$(ch(2), 1, 8)
ch(1) = Mid$(ch(2), 7, 8)
ch(0) = set_protect_code(ch(0))
ch(1) = set_protect_code(ch(1))
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
cal_password = out_string(0) + out_string(2)
End Function

Public Function cal_protect_code(in_ch As String, time%) As String
Dim l!
Dim i%
  l! = l! + Asc(Mid$(in_ch, 1, 1)) Mod 100
  l! = l! * 100 + (Asc(Mid$(in_ch, 2, 1)) Mod 100)
l! = function_for_protect(l!)
For i% = 1 To time% - 1
l! = Abs(l!) Mod 10001
l! = function_for_protect(l!)
Next i%
l! = Abs(l!) Mod 100
cal_protect_code = Trim(str(l!))
If Len(cal_protect_code) = 1 Then
  cal_protect_code = "0" + cal_protect_code
End If
End Function

Public Function function_for_protect(i!) As Long
function_for_protect = -3 * i! * i! + 15 * i! + 5
End Function
Public Sub set_vision_for_student()

End Sub

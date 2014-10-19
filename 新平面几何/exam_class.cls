VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IO_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim exam_wenti_name() As String * 20
Private m_last_record_in_file As Integer
Private m_wenti_record As io_record_type
Public Property Get last_record() As Integer
last_record = m_last_record_in_file
End Property

Public Sub set_exam_list()
Dim t_wenti_name As String
file_record.mark = "DOESSOFT"
Erase exam_wenti_name
m_last_record_in_file = 0
Open App.path & "\example.jak" For Random As #1 Len = Len(m_wenti_record)
m_last_record_in_file = 0
If LOF(1) > 0 Then
Do While EOF(1) <> True
m_last_record_in_file = m_last_record_in_file + 1
Get #1, m_last_record_in_file, m_wenti_record
A% = Asc(Mid$(m_wenti_record.name, 1, 1))
If A% <> 0 Then
If Mid$(m_wenti_record.name, Len(m_wenti_record.name), 1) = "~" Then
'选择题
t_wenti_name = Mid$(m_wenti_record.name, 1, Len(m_wenti_record.name) - 1)
Else
t_wenti_name = m_wenti_record.name
End If
Call change_example_name(t_wenti_name)
ReDim Preserve exam_wenti_name(m_last_record_in_file) As String * 20
 exam_wenti_name(m_last_record_in_file) = t_wenti_name
End If
Loop
End If
Close #1
End Sub
Public Sub add_exam_name_to_list(ob As Object)
Dim i%
ob.Clear
For i% = 1 To m_last_record_in_file
ob.AddItem exam_wenti_name(i%)
Next i%
End Sub
Public Sub input_wenti_from_exam(index As Integer)
Dim i%
Dim t_name$
Call init_conditions(0)
inform.Hide
 regist_data.run_type = 0
Open App.path & "\example.jak" For Random As #1 Len = Len(m_wenti_record)
Get #1, index, m_wenti_record
 Call get_wenti_from_record(m_wenti_record)
Close #1
End Sub
Private Sub Class_Initialize()
  Call set_exam_list
End Sub
Public Sub operate_file(ty As Byte, n%) 'ty操作类型1删除插入2插入,n%插入或删除的位置
Dim i%, j%
'ty=1 delete,ty=2 insert
  Open App.path & "\example.jak" For Random As #1 Len = Len(m_wenti_record)
  Open App.path & "\example.ja_" For Random As #3 Len = Len(m_wenti_record)
  For i% = 1 To n% - 1
     Get #1, i%, m_wenti_record
      Put #3, i%, m_wenti_record
  Next i%
  If ty = 1 Then
   For i% = n% + 1 To m_last_record_in_file
     Get #1, i%, m_wenti_record
      Put #3, i% - 1, m_wenti_record
   Next i%
    For i% = n% To m_last_record_in_file - 1
        exam_wenti_name(i%) = exam_wenti_name(i% + 1)
    Next i%
   exam_form.List1.RemoveItem (n% - 1)
   m_last_record_in_file = m_last_record_in_file - 1
  ElseIf ty = 2 Then
     Put #3, n%, wenti_record
    For i% = n% To m_last_record_in_file
     Get #1, i%, m_wenti_record
      Put #3, i% + 1, m_wenti_record
    Next i%
    m_last_record_in_file = m_last_record_in_file + 1
    ReDim Preserve exam_wenti_name(m_last_record_in_file) As String * 20
    For i% = m_last_record_in_file To n% + 1 Step -1
        exam_wenti_name(i%) = exam_wenti_name(i% - 1)
    Next i%
        exam_wenti_name(n%) = Trim(wenti_record.name)
        Call add_exam_name_to_list(exam_form.List1)
  End If
   Close #1
   Close #3
   Kill (App.path & "\example.jak")
   Name App.path & "\example.ja_" As App.path & "\example.jak"
End Sub
Public Sub save_prove_result(ByVal file_name As String)
'Dim file_name As String
Dim th_no As String
Dim temp_string
Dim i%, n%
i% = InStr(1, file_name, ".", 0)
If i% = 0 Then
file_name = file_name + ".pjh"
Else
file_name = Mid$(file_name, 1, i%) + "pjh"
End If
'On Error GoTo save_prpve_result_error
If path_and_file <> "" Or StrOpenFile <> "" And (save_statue > 0) Then
 'file_name = path_and_file '+ loadresstring_(113)
  Open file_name For Output As #2
  If regist_data.run_type = 0 Then
   Print #2, "SHD_PMJH_V2" '"DSH_PMJH"SHD_PMJH 一版SHD_PMJH_V2本二版
  ElseIf regist_data.run_type = 1 Then
   Print #2, "SHD_VMJH_V2" '"DSH_PMJH"SHD_PMJH 一版SHD_PMJH_V2本二版
  End If
temp_string = ""
'If old_wenti_no = 0 Then
' old_wenti_no = wenti_no
'End If
n% = 0
For i% = 1 To last_conditions.last_cond(1).line_no
 If m_lin(i%).data(0).other_no = i% Then
  n% = n% + 1
 End If
Next i% '共有n条线段
'*********
 th_no = value_to_string(C_display_wenti.m_last_conclusion, 3)
  temp_string = th_no
'*****************
th_no = value_to_string(C_display_wenti.m_last_input_wenti_no, 3)
 temp_string = temp_string + th_no
'*******************
th_no = value_to_string(last_conditions.last_cond(1).point_no, 3)
 temp_string = temp_string + th_no
'*************
th_no = value_to_string(n%, 3) 'last_conditions.last_cond(1).line_no - _
        last_conditions.last_cond(0).line_no, 3)
 temp_string = temp_string + th_no
th_no = value_to_string(last_conditions.last_cond(1).circle_no, 3)
 temp_string = temp_string + th_no
th_no = value_to_string(last_conditions.last_cond(1).con_line_no, 3)
 temp_string = temp_string + th_no
Print #2, temp_string
'***********
For i% = 1 To C_display_wenti.m_last_input_wenti_no
If C_display_wenti.m_no(i%) = -1 Then
   Call C_display_wenti.set_m_condition(0, empty_char, 4)
End If
temp_string = C_display_wenti.save_wenti_string(i%)
Print #2, temp_string
Next i%
For i% = 1 To last_conditions.last_cond(1).point_no
temp_string = from_point_to_string(m_poi(i%).data(0))
Print #2, temp_string
Next i%
'put_wenti_to_problem.last_conditions.last_cond(1).line_no = 0
For i% = 1 To last_conditions.last_cond(1).line_no
 If m_lin(i%).data(0).other_no = i% Then
 temp_string = from_line_to_string(m_lin(i%).data(0))
  Print #2, temp_string
 End If
Next i%
For i% = 1 To last_conditions.last_cond(1).circle_no
  temp_string = from_circle_to_string(m_Circ(i%).data(0))
   Print #2, temp_string
Next i%
'If C_display_wenti.m_last_conclusion > 0 Then
'For i% = C_display_wenti.m_last_conclusion + 1 To C_display_wenti.m_last_input_wenti_no
'      th_no = value_to_string(C_display_wenti.m_wenti_theorem_no(i%), 3)
'       temp_string = Trim(str(C_display_wenti.m_ty(i%))) + _
                 Trim(str(C_display_wenti.m_display_string.item(i%).m_conclusion_or_condition)) + _
                  th_no + from_sqr_no_to_string(C_display_wenti.m_display_string.item(i%).m_string)
'      Print #2, temp_string
'Next i%
'End If
save_statue = 0
End If
Close #2
save_prpve_result_error:
End Sub
Private Function from_circle_to_string(cir As circle_data_type) As String
Dim temp_string
Dim i%
 from_circle_to_string = ""
   temp_string = value_to_string(cir.data0.c_coord.X, 4)
  from_circle_to_string = from_circle_to_string + temp_string
   temp_string = value_to_string(cir.data0.c_coord.Y, 4)
  from_circle_to_string = from_circle_to_string + temp_string
   temp_string = value_to_string(cir.data0.center, 3)
  from_circle_to_string = from_circle_to_string + temp_string
   temp_string = value_to_string(cir.data0.color, 3)
  from_circle_to_string = from_circle_to_string + temp_string
   For i% = 0 To 10
      temp_string = value_to_string(cir.data0.in_point(i%), 3)
  from_circle_to_string = from_circle_to_string + temp_string
   Next i%
  from_circle_to_string = from_circle_to_string + Trim(str(cir.data0.visible))
        temp_string = value_to_string(cir.degree, 3)
  from_circle_to_string = from_circle_to_string + temp_string
     temp_string = value_to_string(cir.parent.element(0).no, 3)
  from_circle_to_string = from_circle_to_string + temp_string
     temp_string = value_to_string(cir.parent.element(1).no, 3)
  from_circle_to_string = from_circle_to_string + temp_string
     temp_string = value_to_string(cir.parent.element(2).no, 3)
  from_circle_to_string = from_circle_to_string + temp_string
End Function


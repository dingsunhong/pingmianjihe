VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "wait_for_aid_point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private wait_for_add_point(255) As wait_for_add_point_type
Private start_no%
Private end_no%
Dim c_data As condition_data_type
'在直接推理过程中,搜集添加辅助线的信息.
Public Sub Class_Initialize() '设置后继指针
 Dim i%
  For i% = 0 To 254 - 1
   wait_for_add_point(i%).next_no = i% + 1
  Next i%
  Me.init
End Sub
Public Sub init() '设置后继指针
  start_no% = 0
  end_no% = 0
End Sub
Public Sub set_wait_for_aid_point(ty As Integer, para() As Integer, ByVal last_para%)
Dim i%, t_start_no%, j%
t_start_no% = start_no%
'整理数据处理
Select Case ty
 Case interset_point_line_line
  If is_line_line_intersect(para(0), para(1), 0, 0, False) > 0 Then
       Exit Sub
  Else
    If para(0) > para(1) Then
     Call exchange_two_integer(para(0), para(1))
    End If
  End If
 Case interset_point_three_line
    If para(0) > para(1) Then
     Call exchange_two_integer(para(0), para(1))
    End If
End Select
Do While t_start_no% <> end_no%
   If ty = wait_for_add_point(t_start_no%).ty Then
    If last_para% = wait_for_add_point(t_start_no%).last_para Then
       For j% = 0 To wait_for_add_point(t_start_no%).last_para - 1
           If wait_for_add_point(t_start_no%).para(j%) <> para(j%) Then
              GoTo set_wait_for_aid_point_mark0
           End If
       Next j%
        Exit Sub
    End If
   End If
set_wait_for_aid_point_mark0:
t_start_no% = wait_for_add_point(t_start_no%).next_no
Loop
For i% = 0 To last_para - 1
 wait_for_add_point(end_no%).para(i%) = para(i%)
Next i%
 wait_for_add_point(end_no%).ty = ty
 wait_for_add_point(end_no%).last_para = last_para
  end_no% = wait_for_add_point(end_no%).next_no
End Sub
Public Function get_wait_for_aid_point() As Byte
Dim C_d As condition_data_type
 If start_no% = end_no% Then
    get_wait_for_aid_point = False
 Else
 Do While start_no% <> end_no%
 Select Case wait_for_add_point(start_no%).ty
 Case verti_
   get_wait_for_aid_point = add_aid_point_for_verti0(wait_for_add_point(start_no%).para(0), _
        wait_for_add_point(start_no%).para(1), wait_for_add_point(start_no%).para(2), 0, C_d, 1)
    If get_wait_for_aid_point > 1 Then
       Exit Function
    End If
 Case interset_point_line_line
   get_wait_for_aid_point = add_interset_point_line_line(wait_for_add_point(start_no%).para(0), _
      wait_for_add_point(start_no%).para(1), wait_for_add_point(start_no%).para(2), _
       wait_for_add_point(start_no%).para(3), wait_for_add_point(start_no%).para(4), _
         wait_for_add_point(start_no%).para(5), wait_for_add_point(start_no%).para(6), _
          C_d)
    If get_wait_for_aid_point > 1 Then
       Exit Function
    End If
 Case interset_point_three_line
    get_wait_for_aid_point = add_interset_point_three_line(wait_for_add_point(start_no%).para(0), _
          wait_for_add_point(start_no%).para(1), wait_for_add_point(start_no%).para(2))
    If get_wait_for_aid_point > 1 Then
       Exit Function
    End If
 Case aid_point_for_circle1
    get_wait_for_aid_point = add_aid_point_for_circle1(wait_for_add_point(start_no%).para(0), _
          wait_for_add_point(start_no%).para(1))
    If get_wait_for_aid_point > 1 Then
       Exit Function
    End If
 Case midpoint_
    get_wait_for_aid_point = add_mid_point(wait_for_add_point(start_no%).para(0), 0, _
         wait_for_add_point(start_no%).para(1), 1)
    If get_wait_for_aid_point > 1 Then
       Exit Function
    End If
 End Select
  start_no% = wait_for_add_point(start_no%).next_no
 Loop
 End If
End Function

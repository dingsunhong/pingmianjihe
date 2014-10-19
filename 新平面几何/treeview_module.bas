Attribute VB_Name = "treeview_module"
Public Sub create_treeview(ByVal date_type As Byte, ByVal date_no%, T_V As TreeView, top_of_treeview%)
Dim temp_record As total_record_type
Dim ik%
Dim ind As String
  last_dis_con_gs(dis_gs_no%) = 0
  Last_hotpoint_of_theorem1 = 0
   Erase Hotpoint_of_theorem1
    last_node_index = 0
  '设置TreeView的几何大小
     T_V.top = top_of_treeview%
      T_V.width = Wenti_form.ScaleWidth - 10
'       If ty = 0 Then
'       T_V.Height = Wenti_form.Picture2.ScaleHeight - top_of_treeview% - 70
'       Else
        T_V.Height = Wenti_form.ScaleHeight - top_of_treeview% - 70
'       End If
T_V.Nodes.Clear
last_node_index = 1
ind = "node" & CStr(last_node_index)
Set nod = T_V.Nodes.Add(, , ind, _
            set_display_string0(date_type, date_no%, False, False, 0, 1, 0, True, False))
ReDim Preserve cond_no(nod.index) As condition_no_type
cond_no(nod.index).ty = date_type
cond_no(nod.index).no = date_no%
Call record_no(date_type, date_no%, temp_record, True, 0, 0)
  If temp_record.record_data.data0.condition_data.condition_no > 0 And temp_record.record_data.data0.condition_data.condition_no < 9 Then
  For i% = 1 To temp_record.record_data.data0.condition_data.condition_no
  Call add_node(T_V, ind, _
       temp_record.record_data.data0.condition_data.condition(i%).ty, temp_record.record_data.data0.condition_data.condition(i%).no)
  Next i%
 End If

End Sub
Public Sub add_node(Tr_v As TreeView, paren As String, _
                       ty As Integer, no%)
Dim nod As Node
Dim i%, gs%, con_no%
Dim ind As String
Dim temp_record As total_record_type
Dim dis_string As String
Dim t_dis_string As String
Dim ch As String
'On Error GoTo add_node_error
last_node_index = last_node_index + 1
ind = "node" & CStr(last_node_index)
t_dis_string = set_display_string0(ty, no%, False, False, 0, 1, 0, True, False)
dis_string = set_display_inform(t_dis_string, ty)
If dis_string = "" Then
 Exit Sub
End If
Set nod = Tr_v.Nodes.Add(paren, tvwChild, ind, _
            dis_string)
ReDim Preserve cond_no(nod.index) As condition_no_type
cond_no(nod.index).ty = ty
cond_no(nod.index).no = no%
If ty = general_string_ Then
   If no% = con_g_s(dis_gs_no%, last_dis_gs(dis_gs_no%)) Then
       last_dis_gs(dis_gs_no%) = last_dis_gs(dis_gs_no%) - 1
        gs% = con_g_s(dis_gs_no%, last_dis_gs(dis_gs_no%))
     Call add_node(Tr_v, ind, general_string_, gs%)
     If general_string(gs%).data(0).record.data0.condition_data.condition_no > 0 And _
                   general_string(gs%).data(0).record.data0.condition_data.condition_no < 9 Then
      For i% = 1 To general_string(gs%).data(0).record.data0.condition_data.condition_no - 1
      Call add_node(Tr_v, ind, general_string(gs%).data(0).record.data0.condition_data.condition(i%).ty, _
                               general_string(gs%).data(0).record.data0.condition_data.condition(i%).no)
      Next i%
      con_no% = general_string(gs%).data(0).record.data0.condition_data.condition_no
      If general_string(gs%).data(0).record.data0.condition_data.condition(con_no%).ty = general_string_ Then
         If general_string(general_string(gs%).data(0).record.data0.condition_data.condition(con_no%).no). _
              data(0).value <> "" Then
                 Call add_node(Tr_v, ind, general_string(gs%).data(0).record.data0.condition_data.condition(con_no%).ty, _
                               general_string(gs%).data(0).record.data0.condition_data.condition(con_no%).no)
   
         End If
      Else
             Call add_node(Tr_v, ind, general_string(gs%).data(0).record.data0.condition_data.condition(con_no%).ty, _
                               general_string(gs%).data(0).record.data0.condition_data.condition(con_no%).no)
      End If
     End If
   Else
   Call record_no(ty, no%, temp_record, True, 0, 0)
   If temp_record.record_data.data0.condition_data.condition_no > 0 And temp_record.record_data.data0.condition_data.condition_no < 9 Then
    For i% = 1 To temp_record.record_data.data0.condition_data.condition_no
     Call add_node(Tr_v, ind, temp_record.record_data.data0.condition_data.condition(i%).ty, _
                               temp_record.record_data.data0.condition_data.condition(i%).no)
    Next i%
   End If
  End If
Else
Call record_no(ty, no%, temp_record, True, 0, 0)
If temp_record.record_data.data0.condition_data.condition_no > 0 And temp_record.record_data.data0.condition_data.condition_no < 9 Then
If temp_record.record_data.data0.condition_data.condition_no < 9 Then
For i% = 1 To temp_record.record_data.data0.condition_data.condition_no
Call add_node(Tr_v, ind, temp_record.record_data.data0.condition_data.condition(i%).ty, _
                               temp_record.record_data.data0.condition_data.condition(i%).no)
Next i%
End If
End If
End If
add_node_error:
End Sub



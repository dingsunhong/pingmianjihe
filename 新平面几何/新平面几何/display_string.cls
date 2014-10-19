VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "display_string"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Type wentitype '???????
 'no As Integer
 'condition(49) As String * 1  '记录输入条件
  'point_no(49)    As Integer  '记录输入点号
  ' inter_set_point_type As Integer '
   ' lin(1 To 4) As Integer
    ' circ(1 To 2) As Integer
     ' poi(1 To 4) As Integer
'End Type
'**********************************
Private mvarm_is_input_number_string As Boolean
'Private mvarm_para(4) As String
Private mvarm_input_id_string As String
Private is_first_new_inupt_type As Boolean
Private WithEvents m_icon As Icon_Class
Attribute m_icon.VB_VarHelpID = -1
Private is_draw_operate As Boolean
Private mvarm_icon_ch As String * 1
Private mvarm_icon_id_string As String
Private mvarm_position(3) As Integer 'mvarm_position(0)起始位置
                                     'mvarm_position(1)结束位置
                                     'mvarm_position(2)总行数
                                     'mvarm_position(3)除"已知",起始位置
Private mvarm_V_position As Integer
Private mvarm_H_position As Integer
Private mvarm_icon_position As Integer
Private mvarm_string(2) As String '显示的字符串
Private mvarm_string_no As Integer '显示的字符串数
'Private mvarm_total_string_no As Integer '总的显示字符串数
Private mvarm_note_string As String
Private mvarm_operate_string As String
Private mvarm_complete_string As String
Private mvarm_complete_set_data As Byte '=0,1,2,4,8
Private m_display_color As Integer
Private m_color(3) As Integer
Private mvarm_last_x_condition As Integer
Private mvarm_inpcond As String '菜单字符串
Private mvarm_wenti_cond_no As Integer '当前输入的条件位置
Private mvarm_wenti_dis_cond_no As Integer '显示位置
Private mvarm_theorem_no As Integer '产生数据使用的定理
Private mvarm_ty As Byte '数据类型
Public m_conclusion_or_condition As Byte '结论,条件
Private mvarm_wenti_data As wentitype
Private mvarm_condition_data As condition_type  '推理数据类型
Private mvarm_init_condition(20) As String * 1
'Private mvarm_no As Integer '菜单字符串编号
Private mvarm_icon_point As Integer '光标在mvarm_string中位置
Private mvarm_is_wenti_complete As Byte '输入完成0 1准备 2完成
Private mvarm_is_picture_complete As Boolean '图形输入完成
Private mvarm_is_input_complete As Boolean
Private mvarm_object As Object '显示的窗口对象
Private T_mvarm_object As Object '
Private mvarm_m_initial_x As Integer
Private mvarm_m_initial_y As Integer
Private set_initial_position As Boolean
Private m_row_space '行宽(与显示对象有关
Private m_relate_data(8) As Integer '产生该数据的相关数据的位置
Private m_modify_cond_no As Integer '修改字符的编号
Private m_modify_cond_position As Integer '修改字符的位置
'热点语句
Public m_is_complete_draw As Byte '0 未画完备1 画完,不初始化画板,2画完,初始化画板
Public hotpoint_X1 As Integer
Public hotpoint_X2 As Integer
Private mvarm_last_input As Integer
Private mvarm_depend_no As Integer
Private backup_string(2) As String
Private backup_m_init_h_position%
Private backup_m_init_v_position%
Private backup_is_select As Byte
Private backup_set_hot_point As Byte
Private backup_brace_mark$
Private backup_note_string$
Public m_solve_string_v_position&
Public m_pass_word_for_teacher As Boolean
Private mvarm_save_string As String
Private object_father As Object
Private mvarm_brace_mark$
Private select_mark(2) As Integer
Private mvarm_set_hot_point As Byte
Private mvarm_is_change As Boolean
Private mvarm_display_or_delete As Byte
Public m_is_icon_display As Boolean
'Public sons As son_data
'Private m_icon As Icon_Class
'设置热点的范围
Public Property Get is_wenti_complete() As Byte
 is_wenti_complete = mvarm_is_wenti_complete
End Property
Public Property Get m_string_no() As Integer
m_string_no = mvarm_string_no
End Property
Public Property Get m_inpcond() As String
 m_inpcond = mvarm_inpcond
End Property
Public Property Get m_wenti_cond_no() As Integer
m_wenti_cond_no = mvarm_wenti_cond_no
End Property
Public Property Get m_string() As String
Dim m%
m% = InStr(1, mvarm_string(0), "{", 0)
If m% = 0 Then
   m_string = mvarm_string(0)
Else
   m_string = Mid$(mvarm_string(0), 1, m% - 1)
   If mvarm_string(1) <> "" Then
      m_string = m_string + ";" + mvarm_string(1)
   End If
   If mvarm_string(2) <> "" Then
      m_string = m_string + ";" + mvarm_string(2)
   End If
End If
End Property
Public Sub set_f_object(ob As Object)
Set object_father = ob
End Sub
Public Property Get m_condition_no() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.m_condition_no
    m_condition_no = mvarm_condition_data.no
End Property

Public Property Get m_condition_ty() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.m_condition_no
    m_condition_ty = mvarm_condition_data.ty
End Property
Public Property Get m_theorem_no() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.m_condition_no
    m_theorem_no = mvarm_theorem_no
End Property
Public Property Get m_ty() As Byte
m_ty = mvarm_ty
End Property
Public Sub Get_wenti_cond()
Dim i%
  wenti_cond0.data = mvarm_wenti_data
End Sub
Public Sub set_m_conclusion_or_condition(ty As Byte)
m_conclusion_or_condition = ty
End Sub
Public Function m_is_wenti_complete() As Boolean
m_is_wenti_complete = mvarm_is_wenti_complete
End Function
Public Function Set_wenti() As Byte
Dim i%
  If wenti_cond0.is_set_data Then
  wenti_cond0.is_set_data = False
  mvarm_wenti_data = wenti_cond0.data
  mvarm_ty = inpcond(mvarm_wenti_data.no).ty
  'If mvarm_wenti_data.no = 8 Then
  '   If mvarm_wenti_data.point_no(12) > 0 Then
  '      If m_poi(m_Circ(mvarm_wenti_data.point_no(12)).data(0).data0.center).data(0).data0.name = "" Or _
  '           m_poi(m_Circ(mvarm_wenti_data.point_no(12)).data(0).data0.center).data(0).data0.name = empty_char Then
  '            mvarm_wenti_data.condition(3) = Chr(13)
  '      End If
  '   End If
  'End If
  Call C_display_wenti.set_m_no(0, wenti_cond0.data.no, wenti_cond0.data.wenti_no)
  'Set_wenti = cond_to_display(1, -1)
  End If
End Function
Public Sub set_m_condition_data(ByVal ty As Integer, ByVal no%)
 mvarm_condition_data.ty = ty
 mvarm_condition_data.no = no%
End Sub
Public Property Get m_condition(ByVal n%) As String
 m_condition = mvarm_wenti_data.condition(n%)
End Property
'Public Property Get para(i%) As String
' para = mvarm_para(i%)
'End Property
Public Property Get m_no() As Integer
 m_no = mvarm_wenti_data.no
End Property
Public Property Get m_no_() As Integer
 m_no_ = mvarm_wenti_data.no_
End Property
Public Property Get m_point_no(ByVal n%) As Integer
 m_point_no = mvarm_wenti_data.point_no(n%)
End Property
Public Property Get m_inner_lin(ByVal n%) As String
 m_inner_lin = mvarm_wenti_data.line_no(n%)
End Property
Public Property Get m_inner_circ(ByVal n%) As String
 m_inner_circ = mvarm_wenti_data.circ(n%)
End Property
Public Property Get m_inner_poi(ByVal n%) As String
 m_inner_poi = mvarm_wenti_data.poi(n%)
End Property
Public Property Get m_inner_point_type() As String
 m_inner_point_type = mvarm_wenti_data.inter_set_point_type
End Property
Public Sub set_m_inner_lin(lin_n%, ByVal n%)
   mvarm_wenti_data.line_no(n%) = lin_n%
End Sub
Public Sub set_m_inner_circ(circ_n%, ByVal n%)
  mvarm_wenti_data.circ(n%) = circ_n%
End Sub
Public Sub set_m_inner_poi(poi_n%, ByVal n%)
   mvarm_wenti_data.poi(n%) = poi_n%
   mvarm_wenti_data.point_no(29) = max(mvarm_wenti_data.point_no(29), poi_n%)
   mvarm_wenti_data.point_no(28) = min(mvarm_wenti_data.point_no(28), poi_n%)
End Sub
Public Sub set_m_inner_point_type(ty_n%)
   mvarm_wenti_data.inter_set_point_type = ty_n%
End Sub

Public Function set_m_condition(ByVal ch$, ByVal n%) As Byte
'将字符ch$安置在n%
Dim i%
   If n% < 0 Then
     n% = mvarm_wenti_cond_no
   End If
   If mvarm_inpcond = "" And mvarm_wenti_data.no > 0 Then
    ch$ = up_low_case(ch$)
   ElseIf mvarm_inpcond <> "" Then '显示语句的模板
    ch$ = up_low_case(ch$)
   End If
   If ch$ <> "_" And ch$ <> global_icon_char Then '
   For i% = 0 To 3
    If inpcond(mvarm_wenti_data.no).taboo(i%).taboo_relation(0) = n% Then
       If mvarm_wenti_data.condition(inpcond(mvarm_wenti_data.no).taboo(i%).taboo_relation(1)) = ch$ Then
          Call display_information(inpcond(mvarm_wenti_data.no).taboo(i%).ty, 4, "输入错误")
            Exit Function
       End If
    End If
   Next i%
  End If
 set_m_condition = set_m_condition_(ch$, n%) ''将字符ch$安置在n%
 Call display_m_input_condi(0, 0, 0)
End Function
Public Function set_m_condition_(ByVal ch$, ByVal n%) As Byte
Dim t_wenti_cond_no%
'向第n句输入字符ch$
If ch$ <> empty_char Then '
If n% >= 0 Then
   t_wenti_cond_no = n% '
    If mvarm_is_input_number_string = False Then
 '      mvarm_wenti_data.condition(n%) = Chr(127)
 '   Else
     If mvarm_wenti_data.condition(n%) = ch$ Then
       Exit Function
     ElseIf n% > 1 Then
      If ch$ = "^" And mvarm_wenti_data.condition(n% - 1) = "^" Then
       Exit Function
      End If
     End If
    End If
Else
   t_wenti_cond_no = mvarm_wenti_cond_no + 1
End If
   If t_wenti_cond_no > mvarm_wenti_cond_no Then
       mvarm_wenti_cond_no = t_wenti_cond_no
   End If
If mvarm_is_input_number_string = True Or (mvarm_is_input_number_string = False And _
       InStr(1, mvarm_string(0), "![\\" + Trim(str(mvarm_wenti_cond_no)) + "\\]") > 0) Then
        mvarm_is_input_number_string = True '输入number_string
         If mvarm_wenti_data.point_no(mvarm_wenti_cond_no) = 0 Then
            mvarm_wenti_data.condition(mvarm_wenti_cond_no) = Chr(127)
            mvarm_wenti_data.point_no(mvarm_wenti_cond_no) = set_number_string(ch$, 0) '建立字符串形式的数字，并将其序号记入point_no
         Else
          If Asc(ch$) = 13 Then '输入数字的结束符
            mvarm_wenti_data.condition(mvarm_wenti_cond_no) = Chr(13)
            mvarm_is_input_number_string = False '输入数字过程结束
          Else '否则修改字符串形式的数字
            Call set_number_string( _
                 Trim(number_string(mvarm_wenti_data.point_no(mvarm_wenti_cond_no)) + ch$), _
                     mvarm_wenti_data.point_no(mvarm_wenti_cond_no))
          End If
         End If
Else
        
        
'End If
'可以不按顺序输入mvarm_wenti_cond_no是输入条件的最大序号
 '***************************************
 If ch$ = "BACKSPACE" Or ch$ = "backspace" Then '
     If t_wenti_cond_no > 0 Then
     mvarm_wenti_data.condition(t_wenti_cond_no - 1) = empty_char
      t_wenti_cond_no = t_wenti_cond_no - 2
       mvarm_wenti_cond_no = mvarm_wenti_cond_no - 2
     Else
       t_wenti_cond_no = -1
       mvarm_wenti_cond_no = -1
       mvarm_condition_data.no = 0
     End If
        mvarm_wenti_dis_cond_no = 0
         mvarm_string(0) = ""
          n% = -1
 Else
       If mvarm_wenti_data.condition(t_wenti_cond_no) = empty_char Then '输入
         mvarm_wenti_dis_cond_no = t_wenti_cond_no
         n% = -1
       Else '修改
         n% = t_wenti_cond_no
       End If
         mvarm_wenti_data.condition(t_wenti_cond_no) = ch$
         If ch$ >= "A" And ch$ <= "Z" Then
          mvarm_wenti_data.point_no(t_wenti_cond_no) = point_number(ch$)
         End If
         If mvarm_wenti_data.point_no(t_wenti_cond_no) > 0 Then
            Call set_point_name(mvarm_wenti_data.point_no(t_wenti_cond_no), mvarm_wenti_data.condition(t_wenti_cond_no))
         End If
 End If
 End If

         set_m_condition_ = cond_to_display(0, n%)  '生成显示字符串
 End If
End Function
Public Function set_m_point_no(point_n%, ByVal n%, is_point_no As Boolean) As Byte
'输入点
  If n% >= 0 And is_point_no Then '输入点
     If mvarm_wenti_data.point_no(n%) > 0 Then '已有点
        Exit Function
     End If
  End If
  If mvarm_wenti_data.point_no(n%) <> point_n% Then '输入不同点,修改点,也可能是新输入点
      mvarm_wenti_data.point_no(n%) = point_n% '修改
    If point_n% > 0 And point_n% <= 100 And is_point_no Then
           Call get_new_char(point_n%) '取点名
      If mvarm_wenti_data.condition(n%) <> m_poi(point_n%).data(0).data0.name Then '名更
       mvarm_wenti_data.condition(n%) = m_poi(point_n%).data(0).data0.name '更名
        mvarm_wenti_cond_no = n% '当前显示的字符序号
          Call set_string0(mvarm_string(0))
      End If
      If mvarm_icon_point > 0 Then '已有显示光标
       If mvarm_icon_point = point_n% Then
           mvarm_wenti_data.condition(n%) = Chr(17)
       End If
      ElseIf mvarm_wenti_data.condition(n%) = global_icon_char And mvarm_icon_ch = "" Then '无显示光标
          mvarm_wenti_data.condition(n%) = Chr(17)
           mvarm_icon_ch = Chr(17)
            mvarm_icon_point = point_n%
      End If
       mvarm_wenti_data.point_no(29) = max(mvarm_wenti_data.point_no(29), point_n%) '记录最大(49)最小(48)点
      If mvarm_wenti_data.point_no(28) = 0 Then
       mvarm_wenti_data.point_no(28) = point_n%
      Else
       mvarm_wenti_data.point_no(28) = min(mvarm_wenti_data.point_no(28), point_n%)
      End If
    End If
  End If
  Call display_m_input_condi(0, False, 1)
  If mvarm_icon_id_string = "" And mvarm_is_input_complete = False Then
   mvarm_is_input_complete = True
   If mvarm_wenti_data.no > 22 Then
        mvarm_complete_set_data = 3
   ElseIf mvarm_complete_set_data <> 1 Then
      mvarm_complete_set_data = mvarm_complete_set_data + 1
   End If
'     Call C_display_wenti.Input_Complete(mvarm_wenti_data.wenti_no, False)
  End If
End Function
Public Sub re_name_for_point(point_n%, icon_point%)
Dim n%, t_mvarm_wenti_cond_no%
Dim n1%, n2% ' n1% point_n%de 位置,n2% icon_point% de 位置
If mvarm_icon_point <> icon_point% Then
mvarm_icon_ch = ""
mvarm_icon_point = icon_point%
End If
n1% = -1
n2% = -1
 For n% = 0 To 20
   If mvarm_wenti_data.point_no(n%) = icon_point% And _
      mvarm_wenti_data.point_no(n%) = point_n% Then
    n1% = n%
    n2% = n%
   ElseIf mvarm_wenti_data.point_no(n%) = point_n% And point_n% > 0 Then
    n1% = n% 'GoTo re_name_point_mark1
   ElseIf mvarm_wenti_data.point_no(n%) = icon_point% And icon_point% > 0 Then
    n2% = n%
   End If
 Next n%
'n% = mvarm_point_condition(n%)
'修改输入
If n1% = -1 And n2% = -1 Then
   Exit Sub
Else
If n1% >= 0 Then
     If mvarm_wenti_data.condition(n1%) <> m_poi(point_n%).data(0).data0.name And _
          mvarm_wenti_data.condition(n1%) <> empty_char Then
       If m_poi(point_n%).data(0).data0.name = "" Then
           mvarm_wenti_data.condition(n1%) = mvarm_init_condition(n1%)
       Else
           mvarm_wenti_data.condition(n1%) = m_poi(point_n%).data(0).data0.name
       End If
     End If
 End If
 If n2% >= 0 Then
                mvarm_wenti_data.condition(n2%) = Chr(17)
                mvarm_icon_ch = Chr(17)
 Else
           mvarm_icon_ch = ""
 End If
           mvarm_is_change = True
            Call cond_to_display0(1, n%)
             Call display_m_input_condi(0, False, 1)
End If
End Sub
Public Function set_m_no(ByVal ty_n As Byte, ByVal no%, w_no%, last_input_type As Byte, _
                               need_draw As Byte, total_display_no As Integer) As Byte
Dim i%, k%
  If mvarm_condition_data.no = no% And mvarm_condition_data.ty = ty_n Then
     Exit Function '已有同样条件
  Else '设置条件的参数
     If ty_n = 0 Then
       mvarm_wenti_data.no = no%
     End If
     '设置新的显示语句
       mvarm_condition_data.no = no%
       mvarm_condition_data.ty = ty_n
       total_display_no = total_display_no + 1
       mvarm_position(0) = total_display_no
       mvarm_position(1) = total_display_no
       mvarm_position(3) = total_display_no
 End If
 If w_no% > 0 Then
     mvarm_wenti_data.wenti_no = w_no% '输入的条件编号
 End If
  If ty_n <> 0 Then ''输入的条件类型
     If last_input_type = 2 Then
      mvarm_ty = 4
     ElseIf last_input_type = 3 Then
      mvarm_ty = 5
     End If
  Else
     mvarm_ty = inpcond(mvarm_wenti_data.no).ty '输入类型
  End If
   If mvarm_ty <> last_input_type Then '输入的条件类型变化,由条件变为结论
        is_first_new_inupt_type = True
       If draw_operate(1, total_display_no) Then
        is_first_new_inupt_type = True
       End If
   End If
   set_m_no = mvarm_ty
  If no > -1000 Then
  m_icon.init
  m_is_complete_draw = need_draw
   Call complete_wenti_input(1, False)  '输入菜单
   For i% = 0 To m_wenti_cond_no
        k% = trans_cond_no_by_relation(i%)
         mvarm_wenti_data.condition(i%) = up_low_case(mvarm_wenti_data.condition(i%))
   Next i%
        Call cond_to_display(0, -1) '生成原始显示字符串
   End If
End Function
Public Sub set_m_no_(ByVal no%)
 mvarm_wenti_data.no_ = no%
End Sub

Private Sub Class_Initialize()
 Call init
End Sub
Public Sub init()
Dim i%
For i% = 0 To 2
 mvarm_string(i%) = ""
 backup_string(i%) = ""
Next i%
 is_draw_operate = False
 mvarm_is_input_number_string = False
 mvarm_note_string = ""
 mvarm_string_no = 1
 mvarm_inpcond = ""
 mvarm_theorem_no = 0
 mvarm_complete_set_data = 0
 mvarm_last_x_condition = 0
 mvarm_ty = 0
 mvarm_wenti_cond_no = 0
 'mvarm_condition_no = 0
 mvarm_input_id_string = "\\0\\"
 mvarm_wenti_dis_cond_no = 0
 m_conclusion_or_condition = 0
 mvarm_condition_data.no = 0
 mvarm_condition_data.ty = 0
 For i% = 0 To 50
 mvarm_wenti_data.condition(i%) = empty_char
 mvarm_wenti_data.point_no(i%) = 0
 Next i%
 m_is_icon_display = False
 mvarm_is_wenti_complete = 0
 mvarm_is_picture_complete = False
 mvarm_is_input_complete = False
 m_modify_cond_no = -1
 m_modify_cond_position = -1
 mvarm_depend_no = 0
 mvarm_last_input = -1
 backup_m_init_h_position% = 0
 backup_m_init_v_position% = 0
 m_solve_string_v_position = 0
 m_pass_word_for_teacher = False
 backup_is_select = 0
 mvarm_save_string = ""
 mvarm_brace_mark$ = ""
 mvarm_icon_ch = ""
 mvarm_icon_point = 0
 mvarm_icon_position = -1
 Set m_icon = Nothing
 Set m_icon = New Icon_Class
End Sub
Public Property Get icon_display() As Boolean
 icon_display = m_icon.m_is_display
End Property

Public Sub set_m_depend_no()
Dim i%
 If mvarm_depend_no = 0 Then
 mvarm_depend_no = 1
 For i% = 0 To 50
   If mvarm_wenti_data.point_no(i%) > 0 And mvarm_wenti_data.condition(i%) >= "A" And _
                mvarm_wenti_data.condition(i%) <= "Z" Then
      Call set_depend_from_point(mvarm_wenti_data.point_no(i%))
   End If
 Next i%
End If
End Sub
Public Property Get m_depend_no()
 m_depend_no = mvarm_depend_no
End Property
Public Sub set_m_string(ByVal brace_mark$, ByVal s1$, ByVal S2$, ByVal s3$, total_display_no As Integer, _
                        ByVal note_str$, ByVal theorem_no%, ByVal concl_or_cond As Integer, set_hot_point As Byte)   '输入
 mvarm_string(0) = s1$
  If S2$ = "" Then
   mvarm_string_no = 1
  Else
   mvarm_string(1) = S2$
  If s3$ = "" Then
   mvarm_string_no = 2
   total_display_no = total_display_no + 1
  Else
   mvarm_string(2) = s3$
   mvarm_string_no = 3
   total_display_no = total_display_no + 2
  End If
  End If
 If theorem_no% <> 0 Then
 mvarm_theorem_no = theorem_no%
 End If
 mvarm_ty = 4 '输出结论
 mvarm_brace_mark$ = brace_mark$
 mvarm_note_string = note_str$
 mvarm_set_hot_point = set_hot_point
 m_conclusion_or_condition = concl_or_cond
   If InStr(1, mvarm_string(0), "\\", 0) = 0 And mvarm_is_wenti_complete < 2 Then
     'mvarm_is_complete = True
      mvarm_is_wenti_complete = 2
       If mvarm_wenti_data.no <> -1000 Then
          Call C_display_wenti.Input_Complete(mvarm_wenti_data.wenti_no, False)
       End If
   End If
 Call display_m_input_condi(0, 1, 1)                                                  'ty =1 新,同样旧
End Sub
Public Sub set_m_inpcond(ByVal s$) '初始输入
 mvarm_inpcond = s$
End Sub
Public Sub set_m_ty(ByVal ty As Byte)
  If mvarm_ty = 0 Then
   mvarm_ty = ty
   Call display_m_input_condi(0, 0, 0)                                                       'ty =1 新,同样旧
  End If
End Sub
Public Function display_m_string_to_ob(string_no%, ob As Object, ByVal color%, ByVal display_or_delete As Byte, _
                ByVal m_position%, ByVal set_hot_point As Byte, is_select As Byte, ty As Byte) As String
Set T_mvarm_object = mvarm_object
Set mvarm_object = ob
display_m_string_to_ob = display_m_string(string_no%, color%, display_or_delete, _
                                 is_select, ty)                                  'ty =1 新,同样旧
Set mvarm_object = T_mvarm_object
End Function
Public Function display_m_string(string_no%, ByVal color%, ByVal display_or_delete As Byte, _
                                     is_select As Byte, ty As Byte) As String                          'ty =1 新,同样旧
 '上下标, m_position%修改字符的位置
'Static pre_char_position As Integer
Dim icon_px%
Dim s1$
Dim ch$
Dim icon_up_down%, j%, k%
Dim id_string As String
Dim t_icon$
Dim p1%, p2%, p3%, no%
Dim tA(2) As String
If display_or_delete = 1 Then
s1$ = Trim(mvarm_string(string_no%))
Else
s1$ = Trim(backup_string(string_no%))
End If
If mvarm_string(0) = "" Then
Call cond_to_display(0, -1)
End If
If display_or_delete = 0 And mvarm_icon_ch <> Chr(17) Then
    Call m_icon.set_m_is_display(False)
End If
'************************************************************************
id_string = mvarm_icon_id_string
If display_or_delete = 1 Then
tA(1) = read_string_from_string(1, s1, "{{", "}}", j%, tA(0), tA(2))
If tA(1) <> "" Then
 Do
  no% = val(Mid$(tA(1), 3, Len(tA(1)) - 4))
  ch$ = mvarm_wenti_data.condition(no%)
  If ch$ = Chr(127) Then
     tA(2) = number_string(mvarm_wenti_data.point_no(no%)) + _
                      Mid$(tA(2), 2, Len(tA(2)) - 1)
  ElseIf ch$ = Chr(13) Then
     tA(2) = number_string(mvarm_wenti_data.point_no(no%)) + _
                      Mid$(tA(2), 3, Len(tA(2)) - 2)
      number_string(mvarm_wenti_data.point_no(no%)) = _
         simple_string(number_string(mvarm_wenti_data.point_no(no%)))
                      ch$ = Chr(127)
  End If
  s1 = tA(0) + tA(2)
  k% = j%
  tA(1) = read_string_from_string(k%, s1, "{{", "}}", j%, tA(0), tA(2))
 Loop Until tA(1) = ""
End If
backup_string(string_no%) = s1$ 'display_m_string
Else
s1$ = backup_string(string_no%)
End If
display_m_string = display_m_text(color, display_or_delete, Trim(s1$), _
                            icon_up_down%, is_select, 0)
mvarm_last_x_condition = max(mvarm_last_x_condition, Int(mvarm_object.CurrentX))
If mvarm_is_wenti_complete = 2 And mvarm_string_no = 1 Then '完成
 Call display_m_remark(display_or_delete)
End If
End Function
               'ty =1 新,同样旧

Private Sub display_m_remark(display_or_delete As Byte)
                            'set_hot_point=0 无注释
                            'set_hot_point=1 有注释
                            'set_hot_point=2 有注释和理由
If backup_note_string <> "" Then
 If (Mid$(backup_note_string, 1, 1) = "(" Or Mid$(backup_note_string, 1, 1) = "(") And _
     (Mid$(backup_note_string, Len(backup_note_string), 1) = ")" Or _
       Mid$(backup_note_string, Len(backup_note_string), 1) = ")") Then
  mvarm_object.CurrentX = mvarm_last_x_condition + m_row_space / 2
  If mvarm_string_no = 1 Or backup_brace_mark = "" Then
    mvarm_object.CurrentY = mvarm_object.CurrentY
  ElseIf mvarm_string_no = 2 Then
    mvarm_object.CurrentY = mvarm_object.CurrentY - m_row_space / 2
  ElseIf mvarm_string_no = 3 Then
    mvarm_object.CurrentY = mvarm_object.CurrentY - m_row_space
  End If
  If backup_set_hot_point = 1 Then
   Call display_m_hotpoint_of_theorem(display_or_delete, backup_note_string)
  ElseIf backup_set_hot_point = 2 Then
   If m_theorem_no <> 0 Then
      mvarm_note_string = Mid$(mvarm_note_string, 1, Len(backup_note_string) - 1) + "," + th_chose(m_theorem_no).TH_name + ")"
   End If
   Call display_m_hotpoint_of_theorem(display_or_delete, backup_note_string)
  End If
 End If
End If
End Sub
Public Sub set_m_char_font(ByVal A$, ByVal i%, ByVal s$)
 Static temp_font As Integer
 If temp_font = middle Then
If A$ = "^" And temp_font = middle Then
    temp_font = up
ElseIf Mid$(A$, 1, 1) = "?" Then

ElseIf A$ = "_" Or (Asc(A$) > 64 And Asc(A$) < 91) Then
  If i% > 2 Then
   If Mid$(s$, i% - 2, 1) = LoadResString_(1410, "") And Mid$(s$, i%, 1) = "(" Then
    up_down = down
   Else
    up_down = middle
   End If
  Else
    up_down = middle
  End If
  ElseIf A$ = "(" Then
  If i% > 1 Then
    If Mid$(s$, i% - 1, 1) = LoadResString_(1410, "") Then
     up_down = down
        Else
    up_down = middle
   End If
  Else
    up_down = middle
  End If

     
 ElseIf A$ = ")" Then
  If i% > 3 Then
   If Mid$(s$, i% - 3, 1) = LoadResString_(1410, "") Then
    up_down = down
      Else
    up_down = middle
   End If
  Else
    up_down = middle
  End If
 
 Else
  up_down = middle
   End If
 ElseIf temp_font = up Then
  up_down = up
  temp_font = middle
End If
End Sub
Public Function from_char_to_m_condition(ByVal n%) As Integer '光标前的字母
Dim i%, k%
Dim ch$
Dim ch1$
from_char_to_m_condition = 0
For i% = 1 To Len(mvarm_string(0))
    ch$ = Mid$(mvarm_string(0), i%, 1) '
    If Len(m_inpcond) >= i% Then
     ch1$ = Mid$(m_inpcond, i%, 1)
       If ch1$ = "_" Or ch1$ = global_icon_char Then
             If n% = i% - 1 Then
               Exit Function
             End If
              from_char_to_m_condition = from_char_to_m_condition + 1
        End If
     Else
       k% = InStr(i%, mvarm_string(0), "_", 0)
         from_char_to_m_condition = from_char_to_m_condition + k%
          Exit Function
     End If
Next i%
End Function
Public Sub display_m_hotpoint_of_theorem(ByVal display_or_delete As Byte, _
                                          ByVal s As String)
Dim i%
Dim A$
Dim t_color As Byte
If display_or_delete = 0 Then
t_color = 15
Else
t_color = 6
End If
Call SetTextColor_(mvarm_object, 6, display_or_delete)
For i% = 1 To Len(s)
A$ = Mid$(s, i%, 1)
If A$ = "(" Then
hotpoint_X1 = mvarm_object.CurrentX - mvarm_H_position
'Hotpoint_of_theorem(Last_hotpoint_of_theorem).cond_no = cond_no
ElseIf A$ = ")" Then
hotpoint_X2 = mvarm_object.CurrentX + 10 - mvarm_H_position
End If
Call display_char(mvarm_object, A$, middle, t_color, display_or_delete, False)
Next i%
End Sub
Public Sub set_m_theorem_no(ByVal no%)
 mvarm_theorem_no = no%
End Sub
Public Function cond_to_display0(ty As Byte, ByVal no%) As Byte
Dim i%, k%
Dim next_icon$
Dim icon$
mvarm_string(0) = mvarm_inpcond
next_icon$ = m_icon.read_icon_from_string(mvarm_string(0), icon$, 0)
Do While next_icon$ <> ""
   If icon$ = global_icon_char Then
      'mvarm_string(0) = Mid$(mvarm_string(0), 1, next_icon% - 1) & mvarm_wenti_data.condition(k%) & _
                     Mid$(mvarm_string(0), next_icon% + 1, Len(mvarm_string(0)) - next_icon%) '替换已输入的字符
      k% = k% + 1
   ElseIf icon$ = Chr(17) Then
      If mvarm_icon_point% <> mvarm_wenti_data.point_no(k%) Then '显示字符串中原有光标位置与新位置不同
      'mvarm_string(0) = Mid$(mvarm_string(0), 1, next_icon% - 1) & mvarm_wenti_data.condition(k%) & _
                     Mid$(mvarm_string(0), next_icon% + 1, Len(mvarm_string(0)) - next_icon%) '替换已输入的字符
      End If
      k% = k% + 1
   Else
      Do While Asc(mvarm_wenti_data.condition(k%)) <> 13
       ' mvarm_string(0) = Mid$(mvarm_string(0), 1, next_icon% - 1) & mvarm_wenti_data.condition(k%) & _
                      Mid$(inpcond(0).inpcond, next_icon%, Len(mvarm_string(0)) - next_icon% + 1)
        k% = k% + 1
      Loop
        'mvarm_string(0) = Mid$(mvarm_string(0), 1, next_icon% - 1) & mvarm_wenti_data.condition(k%) & _
                      Mid$(mvarm_string(0), next_icon% + 1, Len(mvarm_string(0)) - next_icon%)
      k% = k% + 1
   End If
   'next_icon% = m_icon.read_icon_from_wenti(mvarm_inpcond, mvarm_string(0), next_icon% + 1, icon$, 1, False)
Loop
End Function
Public Function cond_to_display(ty As Byte, ByVal no%) As Byte
'no%>= 指定修改某一条件
Dim icon_p%, relation_p%, str_no%
Dim n%, i%, tn%
Dim dis_icon As Boolean
Dim t_ch(1) As String
Dim MDIForm1_Timer1_staute As Boolean
'*************************************************
MDIForm1_Timer1_staute = MDIForm1.Timer1.Enabled
MDIForm1.Timer1.Enabled = False
If mvarm_is_wenti_complete = 2 And ty < 2 Then '完成一条已知语句的输入
          Call m_icon.set_m_is_display(False) '关笔光标
         Exit Function
End If
If mvarm_wenti_data.no = 0 Then '条件类型,无输入
    mvarm_is_wenti_complete = 0
     MDIForm1.Timer1.Enabled = MDIForm1_Timer1_staute
      Exit Function
End If
If ty = 1 Then '先输入条件,再转化显示条件,用于从记录输入问题
  If cond_to_display_(ty) = 255 Then
      cond_to_display = 255
        Call complete_wenti_input(2, True)
  End If
End If
'*********************************
If m_inpcond = "" Then '初始化显示语句
    mvarm_inpcond = inpcond(mvarm_wenti_data.no).inpcond
             mvarm_wenti_dis_cond_no = 0 '将要现示的字位置
End If
If mvarm_string(0) = "" Then
   mvarm_string(0) = m_inpcond '原始语句
End If
'***************************************
If no% >= 0 Then
If ty < 2 Then
      If mvarm_wenti_data.condition(no%) = empty_char Then
            mvarm_string(0) = Mid$(mvarm_string(0), 1, tn% - 1) + _
                Mid$(mvarm_inpcond, n%, 1) + _
                 Mid$(mvarm_string(0), tn% + 1, Len(mvarm_string(0)) - tn%)
       '(mvarm_string(0) 1-tn%) + (mvarm_inpcond,n%)+(mvarm_string(0) tn%-)
       '在tn%插入(mvarm_inpcond,n%)
                        mvarm_wenti_cond_no = mvarm_wenti_cond_no + 1
                        'mvarm_condition_no = no% + 1
      Else
              If m_condition(mvarm_wenti_cond_no) = Chr(13) Then
                  Call replace_string_by_string(mvarm_string(0), "{\\" + Trim(str(mvarm_wenti_cond_no)) + "\\}" + Chr(127) + "_", _
                       "{{" + Trim(str(mvarm_wenti_cond_no)) + "}}" + Chr(127))
              Else
                  Call replace_string_by_string(mvarm_string(0), "[\\" + Trim(str(mvarm_wenti_cond_no)) + "\\]", _
                       "{{" + Trim(str(mvarm_wenti_cond_no)) + "}}" + m_condition(mvarm_wenti_cond_no))
              End If
        '在tn%插入(m_condition(mvarm_wenti_cond_no))
                        If mvarm_is_input_number_string = False Then
                         mvarm_wenti_cond_no = mvarm_wenti_cond_no + 1
                        End If
                        'mvarm_condition_no = no% + 1
     End If
End If
Else 'no% < 0
' icon_p% = m_icon.read_icon_from_wenti(mvarm_inpcond, mvarm_string(0), 1, t_ch(1), ty, dis_icon)
If InStr(1, mvarm_string(0), "\\", 0) = 0 Then
    cond_to_display = 255
      Call complete_wenti_input(2, True)
Else 'to
'********************************************************************************
'生成显示语句
'********************************************************************************
If mvarm_wenti_cond_no >= 0 Then
 Call set_string0(mvarm_string(0))
End If
cond_to_display_out1:
If InStr(1, mvarm_string(0), "\\", 0) = 0 Then
    cond_to_display = 255
      Call complete_wenti_input(2, True)
End If
Call display_m_input_condi(0, 1, 1)
    If ty = 1 And mvarm_is_wenti_complete < 2 Then
     mvarm_string(0) = remove_icon(mvarm_string(0))
       cond_to_display = 255
           Call complete_wenti_input(2, True)
     End If
End If
End If
     If InStr(1, mvarm_string(0), "\\", 0) = 0 Then
     MDIForm1.Timer1.Enabled = MDIForm1_Timer1_staute '完成输入
     Else
     MDIForm1.Timer1.Enabled = True '未完成输入,显示光标
     End If
End Function
Public Function cond_to_display_(ty As Byte) As Byte
'特例集中处理
If m_no = -33 Then
  If m_point_no(8) > 0 And _
     m_Circ(m_point_no(8)).data(0).data0.center = 0 Then
       mvarm_string(0) = LoadResString_(1535, "\\1\\" + m_condition(0) + _
                                              "\\2\\" + set_display_string_circle(m_Circ(m_point_no(8)).data(0).data0) + _
                                              "\\3\\" + m_condition(0) + m_condition(3))
          cond_to_display_ = 255
            Call complete_wenti_input(2, True)
           Exit Function
   End If
ElseIf m_no = -32 Then
 If m_point_no(12) > 0 And _
     m_Circ(m_point_no(12)).data(0).data0.center = 0 Then
     mvarm_string(0) = LoadResString_(1770, _
            "\\1\\" + set_display_string_circle(m_Circ(m_point_no(12)).data(0).data0) + _
            "\\2\\" + m_condition(2) + _
            "\\3\\" + m_condition(3) + m_condition(4) + _
            "\\4\\" + m_condition(5))
     cond_to_display_ = 255
        Call complete_wenti_input(2, True)
      Exit Function
 End If
ElseIf m_no = 7 Then
  If m_point_no(7) > 0 And _
     m_point_no(0) = 0 Then
      mvarm_string(0) = LoadResString_(1775, "\\1\\" + set_display_string_circle(m_Circ(m_point_no(7)).data(0).data0) + _
                                            "\\2\\" + m_condition(2))
       cond_to_display_ = 255
        Call complete_wenti_input(2, True)
        Exit Function
  End If
ElseIf m_no = 8 And m_point_no(3) = 0 Then
        mvarm_string(0) = Mid$(inpcond(m_no).inpcond, 1, Len(inpcond(m_no).inpcond) - 4)
         mvarm_inpcond = m_string
          Exit Function
ElseIf m_no = 11 And m_point_no(12) > 0 Then
  If m_Circ(m_point_no(12)).data(0).data0.center = 0 Then
    mvarm_string(0) = LoadResString_(4130, _
     "\\1\\" + m_condition(0) + _
      "\\2\\" + set_display_line0(m_condition(1) + _
               m_condition(2)) + _
     "\\3\\" + set_display_string_circle( _
         m_Circ(m_point_no(12)).data(0).data0))
       cond_to_display_ = 255
         Call complete_wenti_input(2, True)
        Exit Function
  End If
End If
          mvarm_inpcond = inpcond(mvarm_wenti_data.no).inpcond ', 1, _
           InStr(1, inpcond(mvarm_wenti_data.no).inpcond, " ", 0) - 1)
              mvarm_string(0) = m_inpcond
End Function

Public Function read_string_from_wenti(ByVal st%, en%, str_no%, dis_icon) As String
'从输入读出信息;信息段;
Dim i%, t_st%, t_en%
 str_no% = 0
 t_st% = 0
 If st% = 0 Then
  str_no% = 1
 Else
  Do
   Call read_string_from_wenti(t_st%, t_en%, 0, dis_icon)
    str_no% = str_no% + 1
     t_st% = t_en%
  Loop Until t_en% = st%
 End If
If Asc(m_condition(st%)) > 128 Or Asc(m_condition(st%)) < 10 Or _
        Asc(m_condition(st%)) = 13 Or Asc(m_condition(st%)) = 32 Then
   read_string_from_wenti = ""
    dis_icon = False
    en% = 0
ElseIf m_condition(st%) = ";" Then
 i% = st% + 1
 read_string_from_wenti = ""
 Do
 read_string_from_wenti = read_string_from_wenti + m_condition(i%)
 i% = i% + 1
 Loop Until m_condition(i%) = ";" Or m_condition(i%) = empty_char
 If m_condition(i%) = ";" Then
    en% = i%
 Else
   read_string_from_wenti = m_condition(st%)
   en% = st% + 1
 End If
Else
   read_string_from_wenti = m_condition(st%)
   en% = st% + 1
End If
   If Asc(m_condition(en%)) > 128 Or Asc(m_condition(en%)) < 10 Or _
         Asc(m_condition(en%)) = 13 Or Asc(m_condition(en%)) = 32 Then
          dis_icon = False '输入结束
   End If
End Function
Private Sub trans_m_string_for_value(m%)   'm%~
Dim i%
Dim ch(1) As String
ch(1) = ""
For i% = m% - 1 To 0 Step -1
ch(0) = Mid$(mvarm_string(0), i%, 1)
If ch(0) <> "!" Then
 ch(1) = ch(0) + ch(1)
Else
 GoTo trans_m_string_for_value_mark0
End If
Next i%
trans_m_string_for_value_mark0:
 If ch(1) = "1" Then
   mvarm_string(0) = Mid$(mvarm_string(0), 1, i% - 1) + _
           Mid$(mvarm_string(0), m% + 1, Len(mvarm_string(0)) - m%)
 ElseIf ch(1) = "-1" Then
   If Mid$(mvarm_string(0), i% - 1, 1) = "+" Then
      mvarm_string(0) = Mid$(mvarm_string(0), 1, i% - 2) + _
           "-" + Mid$(mvarm_string(0), m% + 1, Len(mvarm_string(0)) - m%)
   Else
      mvarm_string(0) = Mid$(mvarm_string(0), 1, i% - 2) + _
           "-" + Mid$(mvarm_string(0), m% + 1, Len(mvarm_string(0)) - m%)
   End If
 Else
  Call remove_brace(ch(1))
  If InStr(2, ch(1), "+", 0) = 0 And InStr(2, ch(1), "-", 0) = 0 Then
    If Mid$(ch(1), 1, 1) = "-" Then
     If Mid$(mvarm_string(0), i% - 1, 1) = "+" Then
      mvarm_string(0) = Mid$(mvarm_string(0), 1, i% - 2) + _
          "!" + ch(1) + "~" + Mid$(mvarm_string(0), m% + 1, Len(mvarm_string(0)) - m%)
     End If
    Else
      mvarm_string(0) = Mid$(mvarm_string(0), 1, i% - 1) + _
          "!" + ch(1) + "~" + Mid$(mvarm_string(0), m% + 1, Len(mvarm_string(0)) - m%)
    End If
  Else
      mvarm_string(0) = Mid$(mvarm_string(0), 1, i% - 1) + _
          "!(" + ch(1) + ")~" + Mid$(mvarm_string(0), m% + 1, Len(mvarm_string(0)) - m%)
  End If
End If
End Sub
Private Sub m_cond_relation(ByVal con_no%) '填充相关的条件
Dim t_position(1) As Integer
Dim p%
  If inpcond(mvarm_wenti_data.no).relation(0, 0) = con_no% Then '由数据inpcond读出关系相关
     t_position(0) = inpcond(mvarm_wenti_data.no).relation(0, 1) - con_no%
  End If
  If inpcond(mvarm_wenti_data.no).relation(1, 0) = con_no% Then
     t_position(1) = inpcond(mvarm_wenti_data.no).relation(1, 1) - con_no%
  End If
  '********************************************************************
  If t_position(1) > 0 Then '关系相关成立
  '修改m_inpcond
    Call m_icon.read_icon_from_string(m_inpcond, "", 0)
     mvarm_inpcond = Mid$(m_inpcond, 1, p% - 1) + mvarm_wenti_data.condition(con_no%) + _
                  Mid$(m_inpcond, p% + 1, Len(m_inpcond) - p%)
  '填充相关的条件
    Call m_cond_relation0(con_no%, t_position(1))
  End If
  If t_position(0) > 0 Then
    Call m_icon.read_icon_from_string(m_inpcond, "", 0)
     mvarm_inpcond = Mid$(m_inpcond, 1, p% - 1) + mvarm_wenti_data.condition(con_no%) + _
                  Mid$(m_inpcond, p% + 1, Len(m_inpcond) - p%)
   Call m_cond_relation0(con_no%, t_position(0))
  End If
End Sub
Private Sub m_cond_relation0(ByVal con_no%, ByVal re_con_no%) '填充相关的条件
Dim re_position As Integer
  Call m_icon.read_icon_from_string(mvarm_string(0), "", 0)
   If re_position > 0 Then
      mvarm_string(0) = Mid$(mvarm_string(0), 1, re_position - 1) + _
                        mvarm_wenti_data.condition(con_no%) + _
                          Mid$(mvarm_string(0), re_position + 1, Len(mvarm_string(0)) - re_position)
                     
   End If
End Sub

Public Function up_low_case(ByVal ch$) As String '
Dim po%
If ch$ > "A" And ch$ < "z" Then
If mvarm_string(0) = "" Then
   mvarm_string(0) = mvarm_inpcond
End If
po% = InStr(1, mvarm_string(0), mvarm_input_id_string, 0)
'当前输入位置
 If po% = 0 Then
   up_low_case = ch$
 Else
    If Len(mvarm_string(0)) > po% + Len(mvarm_input_id_string) Then
       If Mid$(mvarm_string(0), po% + Len(mvarm_input_id_string), 2) = "_~" Then
          '输入数值
          up_low_case = LCase(ch$)
       Else
         up_low_case = UCase(ch$)
       End If
    Else
        up_low_case = UCase(ch$)
    End If
  End If
Else
 up_low_case = ch$
End If
End Function
Private Function trans_cond_no_by_relation(no%) As Integer
Dim t_no(1) As Integer
If mvarm_wenti_data.no <> 0 Then
   If inpcond(mvarm_wenti_data.no).relation(0, 0) = no% Then
    t_no(0) = inpcond(mvarm_wenti_data.no).relation(0, 1)
   End If
   If inpcond(mvarm_wenti_data.no).relation(1, 0) = no% Then
    t_no(1) = inpcond(mvarm_wenti_data.no).relation(1, 1)
   End If
   trans_cond_no_by_relation = no%
   If t_no(0) > 0 Then
      If trans_cond_no_by_relation > t_no(0) Then
         trans_cond_no_by_relation = trans_cond_no_by_relation + 1
      End If
   End If
   If t_no(1) > 0 Then
      If trans_cond_no_by_relation > t_no(1) Then
         trans_cond_no_by_relation = trans_cond_no_by_relation + 1
      End If
   End If
End If
End Function

Private Function remove_icon(ByVal ts$) As String
Dim ic$
Dim ip%
remove_icon = Trim(ts$)
ip% = Len(m_icon.read_icon_from_string(remove_icon, ic$, 0))
Do While ip% > 0
If ic$ = "_" Then
   If ip% = 1 Then
      remove_icon = Mid$(remove_icon, 2, Len(remove_icon) - 1)
   ElseIf ip% = Len(remove_icon) Then
      remove_icon = Mid$(remove_icon, 1, Len(remove_icon) - 1)
   Else
      If Mid$(remove_icon, ip% - 1, 1) = "(" And _
           Mid$(remove_icon, ip% + 1, 1) = ")" Then
         remove_icon = Mid$(remove_icon, 1, ip% - 2) + _
                        Mid$(remove_icon, ip% + 2, Len(remove_icon) - ip% - 1)
      Else
         remove_icon = Mid$(remove_icon, 1, ip% - 1) + _
                        Mid$(remove_icon, ip% + 1, Len(remove_icon) - ip%)
      End If
   End If
Else
         remove_icon = Mid$(remove_icon, 1, ip% - 1) + _
                        Mid$(remove_icon, ip% + 1, Len(remove_icon) - ip%)
End If
ip% = Len(m_icon.read_icon_from_string(remove_icon, ic$, 0))

Loop

End Function
Private Sub complete_wenti_input(complete_ty As Byte, dis As Boolean)
  mvarm_is_wenti_complete = complete_ty
   If complete_ty = 2 Then
        Call m_icon.close_icon
'        Call set_number
        Call C_display_wenti.Input_Complete(mvarm_wenti_data.wenti_no, False)
   End If
End Sub
Public Sub set_object(ob As Object, int_x%, int_y%)
 Set mvarm_object = ob
If mvarm_object.ScaleMode = 3 Then
   m_row_space = 20
ElseIf mvarm_object.ScaleMode = 1 Then
   m_row_space = 300
End If

End Sub
Public Function get_MouseDown(dis_no%, Button As Integer, X As Integer, Y As Integer) As Boolean
Dim mouse_h_position%
If Y > mvarm_V_position% And Y < mvarm_V_position% + 20 Then
mouse_h_position% = X - mvarm_H_position
get_MouseDown = True
If Button = 1 Then
   If run_statue >= 5 Then '12.10
    If mouse_h_position% >= hotpoint_X1 And mouse_h_position% <= hotpoint_X2 And mvarm_theorem_no > 0 Then '显示帮助
     MDIForm1.StatusBar1.Panels(1).text = _
      th_chose(mvarm_theorem_no).text
    ElseIf X < hotpoint_X1 Then
      'Call draw_inform(mvarm_condition_data.ty, mvarm_condition_data.no, m_string)
         MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1250, "") '_
    End If
      Call draw_inform(dis_no%, mvarm_condition_data.ty, mvarm_condition_data.no)
      MDIForm1.Timer1.Enabled = True
   Else
    m_modify_cond_position = X
   End If
ElseIf mouse_h_position% >= hotpoint_X1 And mouse_h_position% <= hotpoint_X2 And mvarm_theorem_no > 0 Then '显示帮助
     MDIForm1.StatusBar1.Panels(1).text = _
      th_chose(mvarm_theorem_no).text
End If
End If
End Function
Public Sub display_m_input_condi( _
                         is_select As Byte, ty As Byte, dis As Byte)                       'ty =1 新,同样旧
             'dis=0 ,1, 254 已清屏,不用delete,255 用于原题不动
If set_initial_position = False Then
   mvarm_H_position% = mvarm_m_initial_x 'object_father.currentX_position
   mvarm_V_position% = mvarm_m_initial_y + 20 * (mvarm_position(0) + mvarm_position(2) - 1)
   If mvarm_position(0) + mvarm_position(2) >= mvarm_position(1) Then
      mvarm_position(2) = 0
   End If
If mvarm_ty = 0 Then
 Exit Sub
End If
set_initial_position = True
End If
Call draw_operate(dis, 0)
If is_new_string(mvarm_H_position%, mvarm_V_position%) Or is_select <> backup_is_select _
     Or mvarm_is_change Then
If dis <> 254 Then
 Call display_m_input_condi_(0, is_select, ty, dis)                               'ty =1 新,同样旧
End If
Call Backup(mvarm_H_position%, mvarm_V_position%, is_select)
Call display_m_input_condi_(1, is_select, ty, dis) 'ty =1 新,同样旧
Call print_complete_string(dis)
mvarm_is_change = False
End If
If InStr(1, mvarm_string(0), "\\", 0) = 0 And InStr(1, mvarm_string(0), "_", 0) = 0 Then
mvarm_wenti_cond_no = 0
mvarm_input_id_string = "\\0\\"
End If
End Sub
Public Sub display_m_input_condi_(ByVal display_or_delete As Byte, _
                            is_select As Byte, ty As Byte, dis As Byte)            'ty =1 新,同样旧
 '上下标, m_position%修改字符的位置
'Static pre_char_position As Integer
'　上一次光标的位置,byte=0 ,delete,=1 dis,=2 un_enable
Dim Cu%, Cv%, color%, color_%
Dim i%
Dim c(3) As Integer
Dim co_no%
If display_or_delete = mvarm_display_or_delete Then
   Exit Sub
Else
   mvarm_display_or_delete = display_or_delete
End If
mvarm_last_x_condition = 0
If (m_pass_word_for_teacher = False And dis = 0) And mvarm_wenti_data.no = 0 Then
     Exit Sub
ElseIf mvarm_string(0) = "" Or backup_m_init_v_position% < 0 Then    '无类型
 Exit Sub
ElseIf m_solve_string_v_position& > 0 And _
   backup_m_init_v_position% <= m_solve_string_v_position& And dis = 255 Then
 Exit Sub
End If
     mvarm_object.CurrentX = backup_m_init_h_position%
     mvarm_object.CurrentY = backup_m_init_v_position%
     mvarm_position(1) = mvarm_position(3)
     mvarm_position(2) = mvarm_position(1) - mvarm_position(0) + 1
If m_display_color > 0 Then
   color% = m_display_color
Else
If backup_is_select = 1 Then
 color% = 10
 c(0) = 13
 c(1) = 10 'QBColor(6)
 c(2) = 11
 c(3) = 2
 co_no% = 0
 'color% = 9
ElseIf m_conclusion_or_condition > 0 _
         Or (mvarm_ty > 1 And mvarm_ty < 4) Then
 color% = conclusion_color
 color_% = 3
Else
 color% = condition_color
 color_% = 9
End If
End If
'****************************************************
   Cu% = mvarm_object.CurrentY
    '显示高度
'************************************************************************
If mvarm_string_no = 1 Then
   Call display_m_wenti_no(display_or_delete)
    mvarm_object.CurrentY = Cu% '  恢复高度
 '************************************************************************
      Call display_m_string(0, color%, display_or_delete, _
                                     is_select, ty)
      Call display_m_remark(display_or_delete)
Else
If backup_brace_mark$ = "{" Then
If mvarm_string_no = 2 Then
   mvarm_object.CurrentY = Cu% + m_row_space / 2
   Call display_m_wenti_no(display_or_delete)
     mvarm_object.CurrentY = Cu%
     mvarm_object.CurrentX = mvarm_object.CurrentX + m_row_space / 2
    Call draw_brace_mark(mvarm_object.CurrentX, mvarm_object.CurrentY + m_row_space / 5, _
            mvarm_object.CurrentX + m_row_space / 5, mvarm_object.CurrentY + 2 * m_row_space, color_%)
Else
   mvarm_object.CurrentY = Cu% + m_row_space
   Call display_m_wenti_no(display_or_delete)
     mvarm_object.CurrentY = Cu%
     mvarm_object.CurrentX = mvarm_object.CurrentX + m_row_space / 2
     Call draw_brace_mark(mvarm_object.CurrentX, mvarm_object.CurrentY + m_row_space / 5, _
            mvarm_object.CurrentX + m_row_space / 4, mvarm_object.CurrentY + 3 * m_row_space, color_%)
End If
Cv% = mvarm_object.CurrentX + m_row_space / 2
For i% = 0 To mvarm_string_no - 1
  mvarm_object.CurrentX = Cv%
  mvarm_object.CurrentY = Cu% + i% * m_row_space
  Call display_m_string(i%, color%, display_or_delete, _
                                     is_select, ty)
Next i%
      Call display_m_remark(display_or_delete)
Else
   Cv% = mvarm_object.CurrentX
   Call display_m_wenti_no(display_or_delete)
    mvarm_object.CurrentY = Cu% '  恢复高度
 '************************************************************************
      Call display_m_string(0, color%, display_or_delete, _
                                      is_select, ty)
            Call display_m_remark(display_or_delete)
For i% = 1 To mvarm_string_no - 1
  mvarm_object.CurrentX = Cv% + 2 * m_row_space
  mvarm_object.CurrentY = Cu% + i% * m_row_space
  Call display_m_string(i%, color%, display_or_delete, _
                                       is_select, ty)
Next i%
End If
End If
End Sub


Private Sub display_m_wenti_no(ByVal display_or_delete As Byte)          'ty =1 新,同样旧
Dim Cu%
Dim m_pcoordx0
m_pcoordx0 = mvarm_object.CurrentX
Cu% = mvarm_object.CurrentY
    Call SetTextColor_(mvarm_object, 3, display_or_delete)
      'm_object.CurrentY = Cu% - 2 '升高非汉字
    '显示语句号
'    If wenti_type = 0 Then??
       Call display_char(mvarm_object, Trim(str(mvarm_wenti_data.wenti_no)) & ".", middle, 12, _
             display_or_delete, False)
'    Else '选择题
'     If conclusion_data(0).ty = 0 Then
'       Call display_char(mvarm_object, "A)", middle, 3, display_or_delete, False)
'     ElseIf conclusion_data(0).ty > 0 And conclusion_data(1).ty = 0 Then
'       Call display_char(mvarm_object, "B)", middle, 3, display_or_delete, False)
'     ElseIf conclusion_data(1).ty > 0 And conclusion_data(2).ty = 0 Then
'        Call display_char(mvarm_object, "C)", middle, 3, display_or_delete, False)
'     ElseIf conclusion_data(2).ty > 0 And conclusion_data(3).ty = 0 Then
'       Call display_char(mvarm_object, "D)", middle, 3, display_or_delete, False)
'     Else
'      Call display_char(mvarm_object, Trim(str(mvarm_wenti_no - 4)), middle, 3, display_or_delete, False)
'       mvarm_object.CurrentX = mvarm_object.CurrentX - m_row_space / 5 '+ m_pcoordx0
'      Call display_char(mvarm_object, ".", middle, 3, display_or_delete, False)
'       mvarm_object.CurrentX = mvarm_object.CurrentX + m_row_space / 5 '+ m_pcoordx0
'    End If
'    End If
End Sub
Public Sub set_wenti_no(wenti_n%)
mvarm_wenti_data.wenti_no = wenti_n%
End Sub
Property Get wenti_no() As Integer
wenti_no = mvarm_wenti_data.wenti_no
End Property

Public Sub draw_brace_mark(left As Long, top As Long, right As Long, bottom As Long, color As Integer)
Dim mid_pY&
Dim mid_pX&
Dim width&
'mvarm_object.DrawWidth = 2
mid_pY& = (top + bottom) / 2
mid_pX& = (left + right) / 2
width& = (right - left) / 2
mvarm_object.Line (mid_pX&, top + width&)- _
                  (mid_pX&, mid_pY& - width&), QBColor(color)
mvarm_object.Line (mid_pX&, mid_pY& + width&)- _
                  (mid_pX&, bottom - width&), QBColor(color)
mvarm_object.Circle (right, top + width&), width&, QBColor(color), PI / 2, PI
mvarm_object.Circle (right, bottom - width&), width&, QBColor(color), PI, 3 * PI / 2
mvarm_object.Circle (left, mid_pY& - width&), width&, QBColor(color), 3 * PI / 2, 2 * PI
mvarm_object.Circle (left, mid_pY& + width&), width&, QBColor(color), 0, PI / 2
End Sub
Private Function is_last_input(ByVal n%) As Boolean
Dim i%
If mvarm_last_input >= n% Then
   is_last_input = False
Else
For i% = n% To 29
 If Asc(mvarm_wenti_data.condition(i%)) >= 13 And mvarm_wenti_data.condition(i%) <> "" And _
     mvarm_wenti_data.condition(i%) <> empty_char Then
    is_last_input = False
     mvarm_last_input = i%
      If n% < i% Then
       mvarm_wenti_data.condition(n%) = Chr(13)
      End If
      Exit Function
 End If
Next i%
is_last_input = True
End If
End Function

Private Sub Backup(m_init_h_position%, t_current_y%, is_select As Byte)
Dim i%
For i% = 0 To 2
 backup_string(i%) = mvarm_string(i%)
Next i%
backup_m_init_h_position% = m_init_h_position%
backup_m_init_v_position% = t_current_y%
backup_is_select = is_select
backup_brace_mark$ = mvarm_brace_mark$
backup_set_hot_point = mvarm_set_hot_point
backup_note_string$ = mvarm_note_string$
End Sub

Private Function is_new_string(m_init_h_position%, t_current_y%) As Boolean
 If backup_string(0) <> mvarm_string(0) Then
    is_new_string = True
 Else
   If backup_m_init_h_position% <> m_init_h_position% Or _
       backup_m_init_v_position% <> t_current_y% Then
      is_new_string = True
   ElseIf backup_string(1) <> mvarm_string(1) Or _
        backup_string(2) <> mvarm_string(2) Then
      is_new_string = True
   End If
 End If
End Function
Private Function is_not_last_condition(tn%) As Boolean
Dim i%
For i% = tn% To 29
 If mvarm_wenti_data.condition(i%) <> empty_char And Asc(mvarm_wenti_data.condition(i%)) >= 13 Then
    is_not_last_condition = True
     Exit Function
 End If
Next i%
End Function
Public Function save_string() As String
Dim i%
Dim temp_string As String
If mvarm_save_string <> "" Then
   save_string = mvarm_save_string
Else
If mvarm_wenti_data.no <> 0 Then
 temp_string = value_to_string(mvarm_wenti_data.no, 3)
  mvarm_save_string = temp_string
 temp_string = value_to_string(mvarm_wenti_data.no_, 3)
  mvarm_save_string = mvarm_save_string + temp_string
For i% = 0 To 50
If mvarm_wenti_data.condition(i%) = empty_char Then
 mvarm_save_string = mvarm_save_string + ";"
ElseIf Asc(mvarm_wenti_data.condition(i%)) = 13 Then
 mvarm_save_string = mvarm_save_string + "%"
Else
 mvarm_save_string = mvarm_save_string + mvarm_wenti_data.condition(i%)
End If
temp_string = value_to_string(mvarm_wenti_data.point_no(i%), 5)
  mvarm_save_string = mvarm_save_string + temp_string
Next i%
For i% = 1 To 8
temp_string = value_to_string(mvarm_wenti_data.poi(i%), 5)
  mvarm_save_string = mvarm_save_string + temp_string
Next i%
For i% = 1 To 4
temp_string = value_to_string(mvarm_wenti_data.line_no(i%), 5)
  mvarm_save_string = mvarm_save_string + temp_string
Next i%
For i% = 1 To 4
temp_string = value_to_string(mvarm_wenti_data.circ(i%), 5)
  mvarm_save_string = mvarm_save_string + temp_string
Next i%
temp_string = value_to_string(mvarm_wenti_data.inter_set_point_type, 5)
  mvarm_save_string = mvarm_save_string + temp_string
Else
      temp_string = value_to_string(mvarm_theorem_no, 3)
       temp_string = Trim(str(mvarm_ty)) + _
                Trim(str(m_conclusion_or_condition)) + _
                  temp_string + from_sqr_no_to_string(mvarm_string(0))

End If
 save_string = mvarm_save_string
End If
End Function
Private Sub Class_Terminate()
 mvarm_is_change = True
 Call display_m_input_condi_(0, 0, 0, 1)  'ty =1 新,同样旧
 Set m_icon = Nothing
'ty =1 新,同样旧
End Sub

Public Sub display_icon(display_or_delete As Byte)
 Call m_icon.display_icon(display_or_delete)
End Sub
Private Sub m_icon_IconDisplayChange(ty As Boolean)
 If m_is_icon_display <> ty Then
 m_is_icon_display = ty
 Call C_display_wenti.Event_IconDisplayChange(m_no, ty)
 'Call M_InputEvent.set_input_icon(mvarm_wenti_no, ty)
 End If
End Sub
Public Function is_special_input(m_no%, n_point%, ele1%, ele2%, inter_point_ty%) As Boolean
 If m_no% = mvarm_wenti_data.no And mvarm_wenti_data.poi(1) = n_point% Then
    If m_no% = 9 And is_same_two_point(ele1%, ele2%, _
          mvarm_wenti_data.line_no(1), mvarm_wenti_data.line_no(2)) Then
         is_special_input = True
    ElseIf (m_no = 11 Or m_no = -63) And ele1% = mvarm_wenti_data.line_no(1) And _
                       ele2% = mvarm_wenti_data.circ(1) Then '指定一交圆
         inter_point_ty% = mvarm_wenti_data.inter_set_point_type
         is_special_input = True
    ElseIf (m_no = 13 Or m_no = -67 Or m_no = -66) And _
       is_same_two_point(ele1%, ele2%, _
          mvarm_wenti_data.circ(1), mvarm_wenti_data.circ(2)) Then
         inter_point_ty% = mvarm_wenti_data.inter_set_point_type
         is_special_input = True
    ElseIf m_no = 1 Or m_no% = 7 Or m_no% = -61 Then
         is_special_input = True
    End If
 End If
End Function

Public Function get_number(num_no%) As String
Dim i%, j%, k%
For i% = 1 To 29
    If mvarm_wenti_data.condition(i%) = "!" Then
       k% = k% + 1
       If k% = num_no% Then
          For j% = i% + 1 To 29
               If mvarm_wenti_data.condition(i%) = "~" Then
                  Exit Function
               Else
                  get_number = get_number & mvarm_wenti_data.condition(j%)
               End If
         Next j%
       End If
    End If
Next i%
End Function
Public Function position(i%) As Integer
 position = mvarm_position(i%)
End Function
Public Function Remove_wenti(wenti_no%) As Boolean
If mvarm_wenti_data.wenti_no = wenti_no% Then
   Remove_wenti = True
   Call Class_Terminate
ElseIf mvarm_wenti_data.wenti_no > wenti_no% Then
   Call display_m_input_condi_(0, 0, 0, 1)   'ty =1 新,同样旧
   mvarm_wenti_data.wenti_no = mvarm_wenti_data.wenti_no - 1
   Call display_m_input_condi_(1, 0, 0, 1)   'ty =1 新,同样旧
End If
End Function

Private Function draw_operate(is_draw As Byte, total_display_no As Integer) As Boolean  '显示"已知""求证""证明"
If is_first_new_inupt_type And is_draw_operate = False Then
       is_draw_operate = True
       mvarm_object.CurrentX = mvarm_H_position
       mvarm_object.CurrentY = mvarm_H_position + (mvarm_position(0) + mvarm_position(2) - 1) * 20 ' - 20
draw_operate = True
If mvarm_ty = 1 Then
       Call display_m_text(condition_color, is_draw, LoadResString_(2155, "") + ":", _
                                        mvarm_object.CurrentY, 0, 0)
ElseIf mvarm_ty = 2 Then
       Call display_m_text(condition_color, is_draw, LoadResString_(465, "") + ":", _
                                        mvarm_object.CurrentY, 0, 0)
ElseIf mvarm_ty = 3 Then
       Call display_m_text(condition_color, is_draw, LoadResString_(440, "") + ":", _
                                        mvarm_object.CurrentY, 0, 0)
ElseIf mvarm_ty = 4 Then
       Call display_m_text(condition_color, is_draw, LoadResString_(450, "") + ":", _
                                        mvarm_object.CurrentY, 0, 0)
ElseIf mvarm_ty = 5 Then
       Call display_m_text(condition_color, is_draw, LoadResString_(445, "") + ":", _
                                        mvarm_object.CurrentY, 0, 0)
Else
       draw_operate = False
       Exit Function
End If
       If total_display_no > 0 Then
       total_display_no = total_display_no + 1
        mvarm_position(1) = total_display_no
        mvarm_position(3) = total_display_no
       End If
       mvarm_position(2) = mvarm_position(2) + 1
End If
End Function

Public Sub set_display_init_coordinate(init_x%, init_y%)
If init_x% <> mvarm_m_initial_x Or init_y% <> mvarm_m_initial_y Then
   is_draw_operate = False
Call display_m_input_condi(0, 1, 0)                        'ty =1 新,同样旧
   mvarm_m_initial_x = init_x%
   mvarm_m_initial_y = init_y%
   is_draw_operate = False
Call display_m_input_condi(0, 1, 1)                        'ty =1 新,同样旧
End If
End Sub
Public Sub prove_complete()
mvarm_complete_string = LoadResString_(1700, "") '"证毕"
Call print_complete_string(1)
End Sub
Private Sub print_complete_string(is_draw As Byte)
If mvarm_complete_string = "" Then
   Exit Sub
Else
    mvarm_object.CurrentX = mvarm_H_position% 'object_father.currentX_position
    mvarm_object.CurrentY = mvarm_V_position% + 20 * mvarm_string_no
   Call display_char(mvarm_object, mvarm_complete_string, 0, 0, is_draw, False)
End If
End Sub

Public Sub complete_set_inner_data()
If mvarm_complete_set_data <> 2 Then
 mvarm_complete_set_data = mvarm_complete_set_data + 2
End If
If mvarm_complete_set_data = 3 Then
 Call C_display_wenti.Input_Complete(mvarm_wenti_data.wenti_no, False)
End If
End Sub

Public Sub display_m_string0_(ByVal s1$, ByVal color%, ByVal display_or_delete As Byte, _
                                 is_select As Byte, ty As Byte)                         'ty =1 新,同样旧
Dim t_id_string As String
Dim t_icon$
Dim icon_up_down As Integer
Dim sta_ch As String
Dim ts1$
Dim ts2$
Dim S2$
Dim ts$
Dim icon_px%, icon_py%
Dim p%, n%, p3%, p1%
ts$ = s1$
If ts$ <> "" Then
   ts2$ = read_string_from_string(1, ts$, "[", "]", 0, s1$, ts1$)
        If mvarm_condition_data.ty = 0 And mvarm_condition_data.no > 21 Then
           color = conclusion_color
   Else
           color = condition_color
   End If             'ty =1 新,同样旧
     t_id_string = ""
     Call display_m_text(color%, display_or_delete, s1$, _
                           icon_up_down, is_select, 0)
      If t_id_string = mvarm_icon_id_string Then
       Call m_icon.set_m_icon(t_icon$, icon_up_down, 0, icon_px%, icon_py%, display_or_delete)
      End If
      t_id_string = ""
     Call display_m_text(color%, display_or_delete, ts2$, _
                         icon_up_down, is_select, 0)
     Call display_m_string0_(ts1$, color%, display_or_delete, is_select, ty)
End If
End Sub
Public Sub set_string0(mvarm_s As String) '
Dim t_mvarm_wenti_cond_no As Integer
t_mvarm_wenti_cond_no = mvarm_wenti_cond_no '要修改的字符的序号
Do While mvarm_wenti_data.condition(t_mvarm_wenti_cond_no) <> empty_char And _
      InStr(1, mvarm_s, Trim("\\" + Trim(str(t_mvarm_wenti_cond_no)) + "\\")) > 0 '修改
      'mvarm_wenti_data.condition(t_mvarm_wenti_cond_no)非空，点名
      'Trim("\\" + Trim(str(t_mvarm_wenti_cond_no)) + "\\")非空，插入位置
mvarm_s = LoadResString0_(Trim(mvarm_s), _
                    Trim("\\" + Trim(str(t_mvarm_wenti_cond_no)) + "\\" + _
                        mvarm_wenti_data.condition(t_mvarm_wenti_cond_no)), _
                         1, mvarm_wenti_cond_no, mvarm_wenti_cond_no, backup_string(0))
      '插入
If mvarm_wenti_cond_no = t_mvarm_wenti_cond_no + 1 Then
                   mvarm_input_id_string = "\\" + Trim(str(t_mvarm_wenti_cond_no)) + "\\"
                   mvarm_icon_id_string = m_icon.read_icon_from_string(mvarm_string(0), "", 0)
                    Call m_icon.init
End If
t_mvarm_wenti_cond_no = t_mvarm_wenti_cond_no + 1
Loop 'Until mvarm_wenti_data.condition(t_mvarm_wenti_cond_no) = empty_char  't_mvarm_wenti_cond_no = mvarm_wenti_cond_no
mvarm_wenti_cond_no = t_mvarm_wenti_cond_no
If InStr(1, mvarm_s, Trim("[\\" + Trim(str(t_mvarm_wenti_cond_no)) + "\\]") + "_~") > 0 Then  '下一步输入数字，焦点转移到wenti_form
 Call Wenti_form.Picture1.SetFocus
End If
End Sub
Public Function display_m_text(ByVal color As Byte, _
               ByVal display_or_delete As Byte, ByVal s As String, _
                up_down As Integer, is_select As Byte, is_new_line As Byte, Optional set_icon0 As Boolean = False) As String
                'is_new_line=0不换行 ,is_new_line=1新换行, is_new_line =2,ty =1 新,同样旧
Dim i%, j%, l%, k%, p%, n%
Dim tA(1) As String
Dim tA_(1) As String
Dim control_string As String
Dim t_id_string As String
Dim A$
Dim t_icon$
Dim set_icon1 As Boolean
Dim cu_x1%, cu_y1%, cu_x%, cu_y%
Dim t_color As Byte
Dim up_down_0 As Integer
Dim draw_arrow As Boolean
If s = "" Then
  Exit Function
ElseIf is_new_line > 0 Then '换行
     mvarm_object.CurrentY = mvarm_object.CurrentY + 20
     mvarm_object.CurrentX = 40
     '****************************************************
     If is_new_line = 1 Then
      mvarm_position(1) = mvarm_position(1) + 1
      mvarm_position(2) = mvarm_position(2) + 1
     End If
     '****************************************************
End If
display_m_text = ""
'分解字苻窜
tA_(0) = read_string_from_string_for_print(s, tA(0), tA(1), t_id_string, set_icon1) 's=tA(0)+t_id_string(形式[\\ \\])+tA(1)
If set_icon1 = True Then
 If set_icon0 = False Then
     set_icon0 = True
 Else
     set_icon1 = False
 End If
End If
    If t_id_string <> "" And t_id_string <> change_L Then
       t_color = conclusion_color
'    If mvarm_icon_id_string = t_id_string And mvarm_icon_id_string <> "" Then
'       set_icon = True
'    End If
    Else
       t_color = color
    End If
If tA_(0) = "" Then 'ty =1 新,同样旧
   If t_id_string = change_L Then
      display_m_text = display_m_text & _
           display_m_text(color, display_or_delete, tA(0), up_down, is_select, 0, set_icon0) & _
            t_id_string & _
           display_m_text(color, display_or_delete, tA(1), up_down, is_select, 2, set_icon0)
   Else
    If tA(1) = "" Then '
           A$ = display_char(mvarm_object, tA(0), up_down, color, display_or_delete, 0) '显示字符tA(0)
                 display_m_text = display_m_text & A$
           If A$ = change_L Then
               display_m_text = display_m_text & display_m_text(color, _
                display_or_delete, ByVal s, _
                  up_down, is_select, 1, set_icon0)                                               'ty =1 新,同样旧
           End If
    Else
           display_m_text = display_m_text & _
           display_m_text(color, display_or_delete, tA(0), up_down, is_select, 0, set_icon0) & _
           display_m_text(color, display_or_delete, tA(1), up_down, is_select, 0, set_icon0)
    End If
   End If
Else
tA_(1) = tA_(0)
tA_(1) = Mid$(tA_(1), 2, Len(tA_(1)) - 2)
up_down_0 = 0
draw_arrow = False
If replace_string_by_string(tA_(1), "up", "") Then '
   up_down_0 = 1
ElseIf replace_string_by_string(tA_(1), "down", "") Then
   up_down_0 = -1
   If left(tA_(1), 4) = "([\\" And right(tA_(1), 5) = "\\]_)" Then
       tA_(1) = left(tA_(1), 1) + right(tA_(1), 2)
   End If
End If
If replace_string_by_string(tA_(1), "arrow\\", "") Then
   draw_arrow = True
End If
display_m_text = display_m_text & display_m_text(color, display_or_delete, tA(0), up_down, _
                        is_select, 0, set_icon0)
'用于画向量
cu_x% = mvarm_object.CurrentX
cu_y% = mvarm_object.CurrentY
    A$ = display_char(mvarm_object, tA_(1), up_down_0, color, display_or_delete, False)
  If A$ = change_L Then
             display_m_text = display_m_text & A$
                display_m_text = display_m_text & display_m_text(color, _
                        display_or_delete, tA_(0) & tA(1), _
                           up_down, is_select, 1, set_icon0)                                                'ty =1 新,同样旧
  Else
    If set_icon1 Then
       Call m_icon.set_m_icon(tA_(1), up_down_0, 0, cu_x%, cu_y%, _
                               display_or_delete)
    End If
                    display_m_text = display_m_text & tA_(0) & display_m_text(color, _
                        display_or_delete, tA(1), up_down, is_select, 0, set_icon0)                                                                           'ty =1 新,同样旧
If draw_arrow Then
cu_y1% = mvarm_object.CurrentY
cu_x1% = mvarm_object.CurrentX
mvarm_object.Line (cu_x%, cu_y1% + 6)-(cu_x1%, cu_y1% + 6), t_color
mvarm_object.Line (cu_x1% - 5, cu_y1% + 4)-(cu_x1%, cu_y1% + 6), t_color
mvarm_object.CurrentY = cu_y1%
mvarm_object.CurrentX = cu_x1%
End If
End If
End If
End Function
Private Function display_m_text0(ByVal color As Byte, ByVal up_or_down As Integer, _
               ByVal display_or_delete As Byte, ByVal s As String, Cu%, _
               id_string As String, icon$, up_down As Integer, icon_px%, is_select As Byte) As String
Dim s1$
Dim S2$
Dim s3$
If s = "" Then
   Exit Function
End If
display_m_text0 = ""
S2$ = read_string_from_string(1, s, "!", "~", 0, s1$, s3$)
If S2$ <> "" Then
   Call display_m_text0(color, up_or_down, _
                      display_or_delete, s1$, Cu%, _
                      id_string, icon$, up_down, icon_px%, is_select)
   S2$ = Mid$(S2, 2, Len(S2) - 2)
          Call print_m_string(color, S2, 0, _
               id_string, icon$, up_down, icon_px%, display_or_delete)  '数字
   display_m_text0 = display_m_text0 & S2
   display_m_text0 = display_m_text0 & _
          display_m_text0(color, up_or_down, _
                      display_or_delete, S2$, Cu%, _
                      id_string, icon$, up_down, icon_px%, is_select)
Else
   'If m_icon.read_icon_from_string(A$, icon$, n%) Then
   display_m_text0 = display_m_text0 & _
    display_char(mvarm_object, s, up_or_down, color, display_or_delete, False)
End If
End Function

Private Sub print_m_string(color As Byte, s As String, ty As Byte, _
                              id_string As String, icon$, up_down As Integer, icon_px%, display_or_delete As Byte)
Dim ts$
Dim up_or_down As Integer
If s = "" Then
   Exit Sub
Else
If Mid$(s, Len(s), 1) = "_" Then
   If Len(s) > 1 Then
   ts$ = Mid$(s, 1, Len(s) - 1)
   Call display_char(mvarm_object, ts, up_or_down, color, display_or_delete, False)
   End If
   icon$ = "_"
   id_string = "true"
   up_down = 0
   icon_px% = mvarm_object.CurrentX
   Call display_char(mvarm_object, "_", up_or_down, color, display_or_delete, False)
Else
   Call print_string(mvarm_object, color, s, ty, display_or_delete)
End If
End If
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "display_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public m_display_string As New Collection
'Private m_display_string As New display_string
Private m_yuantibudong As Boolean
Private m_total_display_string_no As Integer
Private m_row_space As Integer '行距
Private m_init_v_position As Integer '
Private m_mody_v_position0 As Integer '显示高度调整
Private m_mody_v_position As Integer '高度调整
Private m_init_h_position As Integer '
Private mvarm_last_condition As Integer
Private m_prove_string As String '求证,求
Private m_solve_string As String '证明,解
Private mvarm_last_conclusion As Integer
Private m_display_complete As Boolean
'Private m_total_display_string_no   As Integer
Private m_pass_word_for_teacher As Boolean
Private m_display_by_step As Boolean
Private m_page_top As Integer
Private m_page_bottum As Integer
Private m_pcoordx0 As Integer
Private m_pcoordy0 As Integer
Private mvarm_chose_wenti_no As Integer '正在操作的语句
Private mvarm_input_statue As Byte '0 输入,1修改,2阅读,3,4,
Public m_page_n As Integer
Private m_object As Object
Private object_me As Object
Private m_wenti_no0 As Integer '显示的起始语距号
Public m_last_input_wenti_no As Integer
Public Last_hotpoint_of_theorem As Integer
Private m_prove_string_v_position&
Private m_solve_string_v_position&
Private last_input_wenti_type As Byte
Private mvarm_condition_for_min_max_value(3) As String
Private mvarm_display_icon As Boolean
Public Sub set_icon_display(ty As Boolean)
Dim i%
For i% = 1 To m_display_string.Count
 If m_display_string.item(i%).icon_display Then
     mvarm_display_icon = True
     MDIForm1.time11_display_type = icon_
      MDIForm1.Timer1.interval = 400
       MDIForm1.Timer1.Enabled = True
      Exit Sub
 End If
Next i%
    MDIForm1.Timer1.Enabled = True
     mvarm_display_icon = False
     MDIForm1.time11_display_type = 0
End Sub
Public Property Get m_icon_display()
m_icon_display = mvarm_display_icon
End Property
Public Sub Set_m_condition_for_min_max_value(no%, st$)
 mvarm_condition_for_min_max_value(no%) = st$
End Sub
Public Property Get m_condition_for_min_max_value(no%) As String
 m_condition_for_min_max_value = mvarm_condition_for_min_max_value(no%)
End Property
Public Property Get m_last_conclusion() As Integer
 m_last_conclusion = mvarm_last_conclusion
End Property
Public Property Get m_last_condition() As Integer
 m_last_condition = mvarm_last_condition
End Property
Public Sub set_m_input_statue(ByVal ty_no As Byte)
mvarm_input_statue = ty_no
End Sub
Public Sub Set_Me_object(ob As Object)
Set object_me = ob
End Sub
Private Sub Class_Initialize()
Set m_display_string = New Collection
 m_row_space = 0 '行距
 m_init_v_position = 0 '
 m_mody_v_position0 = 0 '显示高度调整
 m_mody_v_position = 0 '高度调整
 m_init_h_position = 0
 mvarm_last_condition = 0
 m_prove_string = "" '求证,求
 m_solve_string = "" '证明,解
 mvarm_last_conclusion = 0
 m_display_complete = False
 m_display_by_step = False
 m_page_top = 0
 m_page_bottum = 0
 m_pcoordx0 = 0
 m_pcoordy0 = 0
 mvarm_chose_wenti_no = 0
 mvarm_input_statue = 0 '0 输入,1修改,2阅读,3,4,
 m_page_n = 0
 m_last_input_wenti_no = 0
 Last_hotpoint_of_theorem = 0
 last_input_wenti_type = 0
 m_prove_string_v_position& = 0
 m_solve_string_v_position& = 0
 m_total_display_string_no = 0
 End Sub
Public Sub init(ob As Object, wenti_no0 As Integer)
Call Class_Initialize
Call set_object(ob)
Set m_display_string = Nothing
Set m_display_string = New Collection
m_wenti_no0 = wenti_no0
'Call set_new_item(False)
End Sub
Private Sub set_new_item(dis As Boolean)  'dis=T 显示
Dim temp_string As display_string
Set temp_string = New display_string
If m_display_string.Count > 0 Then
  If m_display_string.item(m_display_string.Count).is_wenti_complete Then
   m_last_input_wenti_no = m_display_string.Count + m_wenti_no0
    If dis Then
     Call input_m_sentences(m_display_string.Count, m_display_string.Count, True, 0)
    End If
   If m_display_string.item(m_display_string.Count).m_is_complete_draw > 0 Then
     Call set_initial_condition(m_display_string.Count, 0, False)
      If m_display_string.item(m_display_string.Count).m_no > 23 Then
       draw_statue = False
      End If
   End If
   'event_statue = ready
   Call C_display_picture.redraw_point(0)
     m_display_string.Add temp_string
   Call m_display_string.item(m_display_string.Count).init
      m_display_string(m_display_string.Count).m_solve_string_v_position& = _
       m_solve_string_v_position&
      m_display_string(m_display_string.Count).m_pass_word_for_teacher = _
       m_pass_word_for_teacher
   Call m_display_string.item(m_display_string.Count).set_object(m_object)
   Call m_display_string.item(m_display_string.Count).set_wenti_no(m_wenti_no0 + _
                  m_display_string.Count)
   Call m_display_string.item(m_display_string.Count).set_f_object(object_me)
  Else
   Set temp_string = Nothing
  End If
Else
      m_display_string.Add temp_string
      m_display_string(m_display_string.Count).m_solve_string_v_position& = _
       m_solve_string_v_position&
      m_display_string(m_display_string.Count).m_pass_word_for_teacher = _
       m_pass_word_for_teacher
   Call m_display_string.item(m_display_string.Count).set_object(m_object)
   Call m_display_string.item(m_display_string.Count).set_wenti_no(m_wenti_no0 + _
            m_display_string.Count)
   Call m_display_string.item(m_display_string.Count).set_f_object(object_me)
End If
End Sub
Public Sub set_m_yuantibudong(yu As Boolean)
 m_yuantibudong = yu
End Sub

Public Function is_same_string(ByVal no%, ByVal s$) As Boolean
Dim ts$

ts$ = m_display_string.item(no%).m_string
If s$ = ts$ Then
   is_same_string = True
Else
   is_same_string = False
End If
End Function
Public Sub input_m_sentences(n1%, n2%, _
                               is_result As Boolean, dis_type As Byte)
'ty=0 ty=1已知 ty=2求证 ty=3 求
Dim n%, step_n%
If n1% > n2% Then
   step_n% = -1
Else
   step_n% = 1
End If
If n1% = 1 And m_display_string.item(n1%).m_no <> 0 And m_pcoordy0 >= 0 Then
m_object.CurrentY = currentY_position(0)
m_object.CurrentX = -m_init_h_position + m_pcoordx0
  Call SetTextColor(m_object.hdc, 0)
   m_object.Print LoadResString_(2155, "") + ":" '已知
End If
For n% = n1% To n2% Step step_n%
  m_object.CurrentY = currentY_position(n%)
  m_object.CurrentX = -m_init_h_position + m_pcoordx0
If m_display_string.item(n%).m_ty = 1 And n% > m_last_conclusion And _
  m_last_conclusion > 0 Then
ElseIf m_display_string.item(n%).m_ty = 2 And m_display_string.item(n1%).m_no <> 0 Then  '结论
If mvarm_last_condition = 0 Then '首次输入结论
If wenti_type = 0 Then
 If n% >= 1 Then
  If n% = 1 Then
  ElseIf m_display_string.item(n% - 1).m_ty = 1 Then  '
  Else
   GoTo input_m_sentences_mark0
  End If
     Call SetTextColor(m_object.hdc, 0)
   If m_prove_string = "" Then
   m_prove_string = LoadResString_(465, "")
   End If
   If m_solve_string = "" Then
   m_solve_string = LoadResString_(450, "")
   End If
    input_p1% = m_object.CurrentY
       m_solve_string = LoadResString_(450, "")
        m_prove_string_v_position& = m_object.CurrentY
        m_object.Print m_prove_string '证求
        mvarm_last_condition = n% - 1
        m_mody_v_position0 = m_mody_v_position0 + 1
  End If
  End If
input_m_sentences_mark0:
 'End If
Else '选择题
 ' If conclusion_data(0).ty = 0 Then
 '  m_object.Print LoadResString_(435,"")
 '   mvarm_last_condition = n% - 1
     'top_sentence_no = n% + 1
 ' End If
End If
ElseIf m_display_string.item(n%).m_ty = 3 And m_display_string.item(n1%).m_no <> 0 Then
Call SetTextColor_(m_object, 0, 1)
 If n% >= 1 Then
  If n% = 1 Then
  ElseIf m_display_string.item(n% - 1).m_ty = 1 And mvarm_last_condition = 0 Then
  Else
   GoTo input_m_setences_mark1
  End If
   Call SetTextColor(m_object.hdc, 0)
    If m_prove_string = "" Then
       m_prove_string = LoadResString_(440, "")
    End If
    If m_solve_string = "" Then
       m_solve_string = LoadResString_(445, "")
    End If
       m_prove_string_v_position& = m_object.CurrentY
    m_object.Print m_prove_string  '求
      m_mody_v_position0 = m_mody_v_position0 + 1
       mvarm_last_condition = n% - 1
   End If
input_m_setences_mark1:
  'End If
ElseIf m_display_string.item(n%).m_ty = 4 Then  'And is_result = False Then
 If n% > 1 Then
  If m_display_string.item(n% - 1).m_ty > 0 Then
   If m_solve_string <> "" Or dis_type = 254 Then
     Call SetTextColor(m_object.hdc, QBColor(0))
      If mvarm_last_conclusion = 0 Then
       m_solve_string_v_position& = m_object.CurrentY
       m_display_string.item(m_display_string.Count).m_solve_string_v_position& = _
           m_solve_string_v_position&
       m_object.Print m_solve_string '证明'解
       m_mody_v_position0 = m_mody_v_position0 + 1
       mvarm_last_conclusion = n% - 1
      End If
   End If
   End If
  End If
End If
'If set_currentY(m_object, n%) = False Then
 '  Exit Sub
'End If
 'm_object.CurrentY = currentY_position(n%)
 If m_solve_string_v_position& > 0 Then
    m_display_string.item(n%).m_solve_string_v_position& = m_solve_string_v_position&
 End If
Call m_display_string.item(n%).display_m_input_condi(True, m_init_h_position, m_init_v_position, 0, 0, dis_type)
Next n%
End Sub
Public Sub set_object(ob As Object)
Dim i%
Set m_object = ob
If m_object.ScaleMode = 3 Then
   m_row_space = 20
   m_page_top% = 90
   m_page_bottum% = 850
ElseIf m_object.ScaleMode = 1 Then
   m_row_space = 300
   m_page_top% = 1400
   m_page_bottum% = 13800
End If
For i% = 1 To m_display_string.Count
 Call m_display_string.item(i%).set_object(m_object, _
                  m_init_h_position, m_init_v_position)
Next i%
End Sub
Public Function set_m_condition(ByVal w_n%, ByVal ch$, n%) As Byte
Dim i%
Dim tn%
tn% = from_wenti_no_to_index0(w_n%, 0) '显示语句的序号
   If n% < 0 Then '未输入条件的位置
     n% = m_display_string.item(tn%).m_wenti_cond_no '最后一个条件的位置
   End If
   If m_display_string.item(tn%).set_m_condition(ch$, n%) = 255 Then
'       If w_n% >= m_display_string.Count Then
'        Call set_new_item(False)
 '      End If
      Call set_initial_condition(w_n%, 0, True)
      If C_display_picture.m_last_draw_point >= last_conditions.last_cond(1).point_no Then
         set_m_condition = 255
      Else
         set_m_condition = 254 '未作图
      End If
   End If
End Function
Public Sub set_m_point_no(ByVal w_n%, ByVal point_n%, ByVal n%, is_point_no As Boolean)
'输入语句中设置点
Dim tn%
tn% = from_wenti_no_to_index0(w_n%, 0) '读出，item的序号
 Call m_display_string.item(tn%).set_m_point_no(point_n%, n%, is_point_no) '显示第n个字符
 '     If w_n% >= m_display_string.Count Then
 '      Call set_new_item(False)
 '     End If
 '  End If
End Sub
Public Sub complete_set_inner_data(w_n%)
Dim tn%
 tn% = from_wenti_no_to_index0(w_n%, 0)
 Call m_display_string.item(tn%).complete_set_inner_data
End Sub
Public Sub re_name_for_point(ByVal point_n%, icon_point%)
Dim i%
'Call C_icon.init
For i% = 1 To m_display_string.Count - 1
  Call m_display_string.item(i%).re_name_for_point(point_n%, icon_point%)
Next i%
End Sub
Public Sub set_m_no(ByVal ty_n As Byte, ByVal no%, w_n%)   'w_n%序号,no%类型
Dim tn%
tn% = from_wenti_no_to_index(w_n%) '读出语句号
If tn% = 1 Or tn% = 0 Then '无输入或只有已知句
   m_total_display_string_no = 0
Else
   m_total_display_string_no = m_display_string.item(tn% - 1).position(1)
End If
   last_input_wenti_type = m_display_string.item(tn%).set_m_no(ty_n, no%, w_n%, last_input_wenti_type, True, _
                                                               m_total_display_string_no)
   mvarm_chose_wenti_no = w_n%
End Sub
Public Sub set_m_no_(ByVal w_n%, ByVal no%)
 Call m_display_string.item(from_wenti_no_to_index0(w_n%, 0)).set_m_no_(no%)
End Sub
Public Sub set_m_inner_poi(w_n%, ByVal poi_no%, pos%)
Dim tn%
tn% = from_wenti_no_to_index0(w_n%, 0)
 Call m_display_string.item(tn%).set_m_inner_poi(poi_no%, pos%)
End Sub
Public Sub set_m_inner_lin(w_n%, ByVal lin_no%, pos%)
Dim tn%
tn% = from_wenti_no_to_index0(w_n%, 0)
 Call m_display_string.item(tn%).set_m_inner_lin(lin_no%, pos%)
End Sub
Public Sub set_m_inner_circ(w_n%, ByVal circ_no%, pos%)
Dim tn%
tn% = from_wenti_no_to_index0(w_n%, 0)
 Call m_display_string.item(tn%).set_m_inner_circ(circ_no%, pos%)
End Sub
Public Sub set_m_inner_point_type(w_n%, ByVal poi_type%)
Dim tn%
tn% = from_wenti_no_to_index0(w_n%, 0)
 Call m_display_string.item(tn%).set_m_inner_point_type(poi_type%)
End Sub

Public Sub display_result(ob As Object, dis As Boolean)  'dis 有许可证.dis_ty=1 整体显示,dis_ty=2 单步显示,
Dim i%, w_n%
If m_pass_word_for_teacher Or dis Then ' = "00000" Or _
       InStr(1, last_conditions.last_cond(1).pass_word_for_teacher, "*", 0) = 0 Then'许可显示
If m_display_by_step = False Then 'da
  ' Call input_m_sentences(1, m_last_conclusion + 1, m_last_input_wenti_no, 1, True, 0)      '显示输入语句
Else
For i% = m_last_conclusion + 1 To m_last_input_wenti_no
Call SetTextColor(m_object.hdc, QBColor(12))
'Call set_currentY(Ob, i%)
m_object.CurrentY = currentY_position(i%)
m_object.CurrentX = m_row_space * 2 - m_init_h_position + m_pcoordx0
m_object.Print LoadResString_(2165, "") '请回车
While code <> 13
    m_object.SetFocus
     DoEvents
    Wend
    code = 0
Call SetTextColor(m_object.hdc, QBColor(15))
'Call set_currentY(Ob, i%)
m_object.CurrentY = currentY_position(i%)
m_object.CurrentX = m_row_space + 2 - m_init_h_position + m_pcoordx0
m_object.Print LoadResString_(2165, "") '请回车
Call Me.display_m_input_condi(i%, 1, 0, 1)
m_object.SetFocus
Beep
Call m_display_string.item(i%).display_m_input_condi(0, True, 0)      '显示输入语句
Next i%
End If
'*****************************************************************************
If InStr(0, m_display_string.item(m_last_input_wenti_no).m_string, LoadResString_(1700, "")) = 0 Then
Call m_display_string.item(m_last_input_wenti_no).set_m_string("", LoadResString_(3955, _
                 "\\1\\" + LoadResString_(1700, "")), "", "", "", 0, 0, w_n%, 1) 'ReDim Preserve display_input_condition(wenti_no) As display_input_cond
'me.m_display_string.item(wenti_no).m_string = LoadResString_(845)
Call m_display_string.item(w_n%).set_m_ty(4)
'Call me.input_m_sentences(ob, 1, w_no%, _
     me.m_last_input_wenti_no, 1, True, 0)
End If
If error_of_wenti = 4 Then
'Call set_currentY(Ob, m_last_input_wenti_no + 1)
m_object.CurrentY = currentY_position(m_last_input_wenti_no + 1)
m_object.CurrentX = m_row_space - m_init_h_position + m_pcoordx0
Call SetTextColor(m_object.hdc, QBColor(12))
m_object.Print LoadResString_(4120, "\\1\\" + str(Abs(error_condition_no)))
                           '"问题的条件不是必要的(条件" + str(Abs(error_condition_no)) + "多余)!"
End If
m_display_complete = True
'operate_step(wenti_no) = last_conditions.last_cond(1).point_no
End If
End Sub
Public Sub click_display_string(ob As Object, ByVal X%, ByVal Y%, modify_or_display As Boolean)
Dim chose_no%
Call set_object(ob)
chose_no% = Y% \ m_row_space
If m_yuantibudong Then
   If chose_no% <= m_last_conclusion + 1 Then
      Exit Sub
   Else
      chose_no% = chose_no% + m_mody_v_position
   End If
Else
   If chose_no% <= m_last_conclusion + m_mody_v_position Then
      Exit Sub
   Else
      chose_no% = chose_no% - m_mody_v_position
   End If
End If
End Sub
Public Sub set_m_pass_word_for_teacher(ty As Boolean) 'ty=true 无需密码
m_pass_word_for_teacher = ty
End Sub

Public Sub set_m_display_by_step(ty As Boolean) 'ty=true 单步
m_display_by_step = ty
End Sub
Public Sub m_change_display(ByVal change_h%, ByVal change_v%)
Dim t_change_h%
Dim t_change_v%
Dim i%

If change_h% <> 0 Then
   t_change_h% = change_h% * m_row_space
   If m_init_h_position <> t_change_h% Then
      m_init_h_position = t_change_h%
   Else
    Exit Sub
   End If
ElseIf change_v% <> 0 Then
   t_change_v% = -change_h% * m_row_space
   If m_init_v_position <> t_change_v% Then
      m_init_v_position = t_change_v%
   Else
    Exit Sub
   End If
Else
 Exit Sub
End If
Call m_object.Cls
For i% = 1 To m_display_string.Count
    Call m_display_string.item(i%).set_display_init_coordinate(m_init_h_position, m_init_v_position)
Next i%
End Sub
Public Sub set_oringinal_point(ByVal X%, ByVal Y%)
m_pcoordx0 = X%
m_pcoordy0 = Y%
End Sub
Public Sub m_Print_wenti(ob As Object, ty As Byte, ty1 As Byte, page_n%, is_continue As Boolean)
Dim i% 'ty=0 显示,ty=1 打印, ty1 线宽
Dim s_x%, s_y%, s_no%
If Me.m_last_input_wenti_no <= 0 Then
 Exit Sub
End If
 Call set_object(ob)
If ty = 1 Then
 printrx& = m_object.TwipsPerPixelX
 printry& = m_object.TwipsPerPixelY
 If page_n% > 0 Then
           Printer.NewPage
 End If
Else
  m_object.Cls
End If
  s_x% = 2000
If ty = 0 Then
 If page_n% = 0 Then
    m_mody_v_position = 0
    m_mody_v_position0 = 0
    m_init_h_position = 0
     s_y% = 19 * m_row_space '5800
 Else
     s_y% = 5 * m_row_space '1600
    m_mody_v_position = 21 + 40 * (page_n% - 1) + m_mody_v_position0
 End If
  Call Me.set_oringinal_point(s_x%, s_y%)
  Call m_P_page(ob, ty, page_n%)
 If page_n% = 0 Then
      Call C_display_picture.print_picture(ob, p_x% + m_row_space, p_y%, ty, ty1)
  If Me.m_last_input_wenti_no > 20 Then
     Call Me.input_m_sentences(1, _
             20, False, 0)
      If is_continue Then
       Call m_Print_wenti(ob, ty, ty1, 1, is_continue)
      End If
  Else
     Call Me.input_m_sentences(1, _
            Me.m_last_input_wenti_no, False, 0)
  End If
 Else
  If Me.m_last_input_wenti_no > 20 + 40 * page_n% Then
     Call Me.input_m_sentences(20 + 40 * (page_n% - 1) + 1, _
             20 + 40 * page_n%, False, 0)
     If is_continue Then
      Call m_Print_wenti(ob, ty, ty1, page_no% + 1, is_continue)
     End If
  Else
      Call Me.input_m_sentences(20 + 34 * (page_n% - 1) + 1, _
             Me.m_last_input_wenti_no, False, 0)
 End If
 End If
ElseIf ty = 1 Then '
s_no% = 0
 If page_n% = 0 Then
  Call C_display_picture.print_picture(ob, p_x%, p_y%, ty, ty1)
 End If
For i% = 1 To Me.m_last_input_wenti_no
'Call c_display_Print_string(Ob, i%, s_x%, s_y%, s_no%,)
If m_object.CurrentY > m_page_bottum% - 800 Then
    Printer.NewPage
     s_no% = i%
      page_n% = page_n% + 1
     If page_n% <= page_no% Then
      Call m_P_page(ob, ty, page_n%)
       s_y% = 1100
     End If
End If
Next i%
End If
If last_p_theorem > 0 Then
m_object.CurrentX = 2000
m_object.Print LoadResString_(4125, "");
 For i% = 1 To last_p_theorem
 m_object.CurrentX = 2000
 m_object.Print th_chose(p_theorem(i%)).TH_name + ":" + _
      th_chose(p_theorem(i%)).text
  If m_object.CurrentY > m_page_bottum% - 800 Then
    page_no% = page_no% + 1
   If ty = 1 Then
    Printer.NewPage
   Else
     m_page_bottum% = m_page_bottum% + 16000
     m_page_top% = m_page_top% + 16000
   End If
      Call m_P_page(ob, ty, page_no%)
   End If
Next i%
End If
If ty = 1 Then
 Printer.EndDoc
End If
End Sub
Private Sub m_P_page(ob As Object, ty As Byte, page_n%)
 If ty = 0 Then
      m_object.DrawWidth = 1
 Else
      m_object.DrawWidth = 2
 End If
      m_object.FontSize = 8
      m_object.CurrentX = m_pcoordx0
      m_object.CurrentY = m_page_top% - m_row_space   '200 '1100
      m_object.Print page_note_string; '页眉"《DSH-平面几何》解题王";
      m_object.CurrentX = 28 * m_row_space '8500
      m_page_n = page_n% + 1
      m_object.Print str(page_n% + 1)
      '上界线
      m_object.Line (m_pcoordx0 - m_row_space, m_page_top%)-(m_pcoordx0 + 24 * m_row_space, m_page_top%), QBColor(9)
      '下界线
      m_object.Line (m_pcoordx0 - m_row_space, m_page_bottum%)-(m_pcoordx0 + 24 * m_row_space, m_page_bottum%), QBColor(9)
      m_object.FontSize = 14
      m_object.Print "         ";
End Sub
Private Function get_wenti_no_row(ByVal row_no%, ByVal cond_no2%, ByVal cond_no1%, ByVal row_no0%) As Integer
Dim i%
If cond_no1% > 0 Then
For i% = cond_no1% To cond_no2%
 row_no0% = row_no0% + m_display_string.item(i%).m_string_no
 If row_no% < row_no0% Then
     get_wenti_no_row = i%
      Exit Function
 End If
Next i%
End If
End Function
Public Sub Set_wenti()
Dim tn%
If wenti_cond0.is_set_data Then
tn% = from_wenti_no_to_index(wenti_cond0.data.wenti_no)
m_last_input_wenti_no = m_display_string.Count
Call m_display_string.item(tn%).Set_wenti
End If
End Sub
Public Sub input_char0(ch As String)
 Call Me.set_m_condition(0, ch$, -1)
 Call set_new_item(True)
End Sub
Public Sub input_char(ch As String)
Dim i As Integer
Dim input_po  As Integer
Dim m%, l%, tp%
Dim f As Integer '用于验证输入是否是新点
Dim t1 As Boolean
Dim t2 As Byte
If event_statue = wait_for_draw_point Then
 event_statue = wait_for_input_char
End If
MDIForm1.Timer1.Enabled = False
 Wenti_form.Label1.visible = False
 If ch = "" Then
 '重操作空输入
 Wenti_form.SetFocus
 Exit Sub
 ElseIf ch = "!" Then
    ch = Chr(12) '","
 End If
'********************************
If event_statue = wait_for_modify_sentence Then
'修改状态
 If ch = "y" Or ch = "Y" Then
   MDIForm1.Timer1.Enabled = False
  'If m_display_string.item(modify_wenti_no).m_no < 23 Then
   '  Call input_m_sentences(0, modify_wenti_no, modify_wenti_no, 1, False, 0)
  'Else
   '  Call input_m_sentences(0, modify_wenti_no, modify_wenti_no + 1, 1, False, 0)
  'End If
   l% = get_last_used_char(modify_wenti_no, 10)
     For i% = 0 To 10
     If m_display_string.item(modify_wenti_no).m_condition(i%) >= "A" And _
          m_display_string.item(modify_wenti_no).m_condition(i%) <= "Z" Then
         If char_number(m_display_string.item(modify_wenti_no).m_condition(i%)) > l% Then
          Call delete_point_name(modify_wenti_no, i%)
         End If
     End If
    Next i%
    Call init_input_cond(modify_wenti_no)
     MDIForm1.Inputcond.Enabled = True
      MDIForm1.conclusion.Enabled = True
      Wenti_form.Label1.visible = False
       event_statue = wait_for_input_sentence
 ElseIf ch = "n" Or ch = "N" Then
      Wenti_form.Label1.visible = False
       If m_display_string.item(modify_wenti_no).m_no < 23 Then
   Call input_m_sentences(modify_wenti_no, modify_wenti_no, False, 0)
  Else
     Call input_m_sentences(modify_wenti_no, modify_wenti_no + 1, False, 0)
 End If
        MDIForm1.Inputcond.Enabled = True
      MDIForm1.conclusion.Enabled = True
    
    event_statue = wait_for_input_sentence
 End If
ElseIf event_statue = wait_for_input_char Or event_statue = _
    wait_for_modify_char Or event_statue = input_char_again Then
 '输入状态
    If (condition_no = -1 And event_statue <> wait_for_modify_char) Or _
         (modify_condition_no = -1 And event_statue = wait_for_modify_char) _
           Then
       Wenti_form.SetFocus
        Exit Sub
    End If
    If event_statue = wait_for_modify_char Then
     event_statue = wait_for_input_char '输入字母
      condition_no = modify_condition_no '字母位置
       modify_condition_no = -1
        l% = get_last_used_char(modify_wenti_no, condition_no - 1)
         If char_number(m_display_string.item(modify_wenti_no).m_condition(condition_no)) > l% Then
                   Call delete_point_name(modify_wenti_no, condition_no)
         End If
           Call m_display_string.item(modify_wenti_no).set_m_condition(empty_char, condition_no)
    End If
     If event_statue = input_char_again Then '保留输入位置
      tp% = condition_no
     End If
     'Call display_m_input_condi(modify_wenti_no, 0, 0, 0)
    If event_statue = input_char_again Then
     MDIForm1.Timer1.Enabled = True
      condition_no = tp%
    End If
   Call m_display_string.item(modify_wenti_no).set_m_condition(ch, condition_no)
  ' Call set_up_case(modify_wenti_no, condition_no) '输入大字母
      ch = m_display_string.item(modify_wenti_no).m_condition(condition_no)
    If Asc(ch) > 63 And Asc(ch) < 91 Then
     Call m_display_string.item(modify_wenti_no).set_m_point_no(point_number(ch), condition_no, True)
      t1 = get_input_info(modify_wenti_no, condition_no)
    Else
     Call next_char(0, ch, 0, 0)
    End If
     t2 = m_display_string.item(modify_wenti_no).cond_to_display(0)
     If t2 = 1 Then
     GoTo input_char_mark0
     End If
     't1=true输入有错　　t2＝true 输完
    If t1 Then
    '  'MsgBox input_char_info, vbOKOnly + 65536
    '   Wenti_form.Label1.Caption = input_char_info
    '   Call dispaly_input_inform
    event_statue = input_char_again
    Else
    event_statue = wait_for_input_char
    If Asc(ch) > 63 And Asc(ch) < 91 Then
     tp% = is_used_char(modify_wenti_no, condition_no)
     If tp% = 0 Then ', is_used_char(char, last_used_char)) = 0 Then
       input_last_point = input_last_point + 1
      Else
      Call m_display_string.item(modify_wenti_no%).set_m_point_no(tp%, condition_no, True)
      End If
    End If
      event_statue = wait_for_input_char
    End If
     If t2 And t1 = False Then
      MDIForm1.Timer1.Enabled = False
     End If
       Call display_m_input_condi(modify_wenti_no, _
            0, 0, 0)
 End If
'End If
If event_statue = wait_for_modify_char And MDIForm1.Timer1.Enabled Then
             '　修改输入
        MDIForm1.Timer1.Enabled = False
         If modify_wenti_no < last_condition Then
          'Call display_m_input_condi(0, _
            modify_wenti_no, 0, 0, 0)
         Else
          'Call display_m_input_condi(0, _
            modify_wenti_no, 0, 0, 0)
         End If '显示修改语句
         second_part$ = ch
          Call m_display_string.item(modify_wenti_no).set_m_string( _
           "", first_part$, second_part, last_part$, "", condition_color, 0)
                 If modify_wenti_no < last_condition Then
          Call display_m_input_condi(modify_wenti_no, 0, 0, 0)
         Else
          Call display_m_input_condi(modify_wenti_no, 0, 0, 0)
         End If '显示修改后语句
    Wenti_form.SetFocus
    Exit Sub
'**********************************************
         'm% = m% / 2    '第一部分的字符数
'**********************************************
         If code = 38 Then
           second_part$ = "^"
          ElseIf code = 40 Then
           second_part$ = "!"
          Else
           second_part$ = ch
          End If
          Wenti_form.SetFocus
         Exit Sub
'**********************************************
          Call relation_char(modify_wenti_no, m% + 1)
           '  修改相关的字符
           modify_statue = no_modify
           ' 关闭修改操作
          MDIForm1.Timer1.Enabled = False
          If modify_wenti_no < last_condition Then
           Call display_m_input_condi(modify_wenti_no, 0, 0, 0)
          Else
           Call display_m_input_condi(modify_wenti_no, 0, 0, 0)
          End If
End If
If Asc(ch) = &HD& And event_statue = modify Then  'Wenti_form.Timer2.Enabled <> 0 And KeyAscii = &HD& Then     '回车
 '删除一句
    'MDIForm1.Timer2.Enabled = 0
          modify_statue = no_modify
       If modify_wenti_no < last_condition Then
        'Call display_m_input_condi(0, _
         modify_wenti_no, 0, 0, 0)
         For i = modify_wenti_no + 1 To last_condition - 1
          Call display_m_input_condi(i, _
                0, 0, 0)
'            Call display_m_input_condi(0, i, _
                 0, 0, 0)
         Call m_display_string.item(i).Get_wenti_cond
         Call m_display_string.item(i - 1).set_wenti_cond
          Next i
'***************************************************************
'****************************************************************
         Call SetTextColor(m_object.hdc, QBColor(0))
          m_object.CurrentY = 20 * last_condition
           m_object.CurrentX = 0
            solve_problem_type = 0
            m_object.Print LoadResString_(465, "")
        Call SetTextColor(m_object.hdc, QBColor(15))
            m_object.Print LoadResString_(465, "")
'********************************************************
        input_type = input_for_conclusion
        For i = last_condition To m_last_input_wenti_no
         Call display_m_input_condi(i, _
                 0, 0, 0)
           'Call display_m_input_condi(0, i, _
                 0, 0, 0)
           Call m_display_string.item(i).Get_wenti_cond
           Call m_display_string.item(i - 1).set_wenti_cond
         ' Set m_display_string.item(i - 1) = m_display_string.item(i)
        Next i
'**********************************************************
   Else
    '**********************************************
    input_type = input_for_conclusion
    'Call display_m_input_condi(0, _
       modify_wenti_no, 0, 0, 0)
           For i = modify_wenti_no + 1 To m_last_input_wenti_no
            Call display_m_input_condi(i, _
                0, 0, 0)
              'Call display_m_input_condi(0, i, _
                  0, 0, 0)
               Call m_display_string.item(i).Get_wenti_cond
               Call m_display_string.item(i - 1).set_wenti_cond
           Next i
     End If
      If modify_wenti_no < last_condition Then
         last_condition = last_condition - 1
         End If
        modify_statue = no_modify
      End If
If (t2 And t1 = False) Then  '_
input_char_mark0:
 condition_no = -1
 If run_statue < 3 Then '12.10
 MDIForm1.Inputcond.Enabled = True
 MDIForm1.conclusion.Enabled = True
 MDIForm1.Toolbar1.Buttons(21).Image = 33
 ElseIf run_statue = 5 Then '12.10
 'MDIForm1.method3.Enabled = True
 ElseIf run_statue = 4 Then '12.10
 MDIForm1.method2.Enabled = True
 End If
 'Or event_statue = wait_for_modify_char Then '输完一句且没
 If list_type_for_input = input_condition_statue And set_or_prove < 2 Then
 If m_display_string.item(modify_wenti_no).m_no < 23 Then
 Call set_condition(modify_wenti_no) '
  If event_statue <> wait_for_modify_char Then
  For i% = draw_wenti_no To m_last_input_wenti_no
   event_statue = wait_for_draw_picture
    Call draw_picture(i%, 0, False)
     If event_statue = wait_for_input_char Then
      Wenti_form.SetFocus
      Exit Sub
     End If
  Next i%
  End If
 Else
  For i% = draw_wenti_no To m_last_input_wenti_no
   event_statue = wait_for_draw_picture
   Call draw_picture(i%, 0, False)
     If event_statue = wait_for_input_char Then
      Wenti_form.SetFocus
      Exit Sub
     End If
   Next i%
    event_statue = complete_input
     If MDIForm1.method0.Checked Then
      run_statue = 2 '12.10
     ElseIf MDIForm1.method1.Checked Then
      run_statue = 3 '12.10
     ElseIf MDIForm1.method2.Checked Then
      run_statue = 4 '.12.10
     Else
      run_statue = 0
     End If
 End If
ElseIf set_or_prove = 2 Then 'list_type_for_input = input_prove_by_hand Then
 If record0.data0.condition_data.condition_no = 254 Then
  Call display_prove_inform(modify_wenti_no, delete)
   'Call display_m_input_condi(0, modify_wenti_no, 0, 0, 0)
    record0.data0.condition_data.condition_no = 255
 End If
  Call draw_picture(modify_wenti_no, 0, True)
ElseIf list_type_for_input = input_add_point Then
 Call draw_picture(m_last_input_wenti_no, 0, False)
 Call display_run(2)
End If
End If
m_object.SetFocus
 'End If
End Sub
Public Sub m_input_char(ob As Object, ch As String)
Dim tn%, point_no%
Dim ty As Byte
If event_statue = wait_for_draw_point Then '
 event_statue = wait_for_input_char
End If
m_last_input_wenti_no = m_display_string.Count
If m_display_string.item(m_display_string.Count).m_no <> 0 Then
    tn% = -1
       ty = set_m_condition(m_display_string.Count, ch, tn%)
End If
If ty = 0 Then
 m_object.SetFocus '继续输入字符
ElseIf ty = 254 Or _
       (ty = 255 And m_display_string.item(m_last_input_wenti_no).m_is_complete_draw) Then '输入字符结束
 Call draw_picture(m_last_input_wenti_no, 0, False)
End If
End Sub
Public Function get_MouseDown(Button As Integer, X As Integer, Y As Integer, w_no%) As Boolean
Dim i%
w_no% = 0
For i% = 1 To m_display_string.Count
If m_display_string.item(i%).get_MouseDown(i%, Button, X, Y) Then
'鼠标选中的语句
w_no% = i%
Exit Function
End If
Next i%
      MDIForm1.StatusBar1.Panels(1).text = "" '_
      Call draw_inform(0, 0, 0)
      MDIForm1.Timer1.Enabled = False
'Exit Function
'get_MouseDown_mark1:
'If w_no% > 0 And m_last_input_wenti_no > 0 Then '选中
'  get_MouseDown = True
'  If mvarm_chose_wenti_no <> w_no% Then
'  If mvarm_chose_wenti_no > 0 Then '以前选中的
'    Call m_display_string.item(mvarm_chose_wenti_no).display_m_input_condi(False, 0, 0, _
                   0, 0, 1) '恢复
'  End If
'  mvarm_chose_wenti_no = w_no%
'    Call m_display_string.item(mvarm_chose_wenti_no).get_MouseDown(Button, X)
'  End If
'ElseIf w_no% = 0 And mvarm_chose_wenti_no > 0 Then '未选中
'  Call m_display_string.item(mvarm_chose_wenti_no).display_m_input_condi(0, 0, 1) '254 已清屏,不用delete,255 用于原题不动 0, _
                   0, 0, 1) '恢复
'  mvarm_chose_wenti_no = 0
'End If
End Function
Public Sub display_m_input_condi(ByVal cond_no As Integer, ByVal set_hot_point As Byte, is_select As Byte, _
         ty As Byte) 'ty =1 新,同样旧
'Call set_currentY(m_object, cond_no%)
'm_object.CurrentY = currentY_position(cond_no)
Call m_display_string.item(cond_no).display_m_input_condi( _
               is_select, ty, 1)               'ty =1 新,同样旧
End Sub
Public Sub set_m_conclusion_or_condition(w_n%, ty As Byte)
 If w_n% > m_last_input_wenti_no Then
 Do
    Call set_new_item(True)
 Loop Until w_n% = m_last_input_wenti_no
ElseIf m_last_input_wenti_no = 0 Then
    Call set_new_item(False)
End If
If w_n% = 0 Then
 If m_last_input_wenti_no > 0 Then
 If m_display_string.item(m_last_input_wenti_no).is_wenti_complete Then
 w_n% = m_last_input_wenti_no + 1
 Else
 w_n% = m_last_input_wenti_no
 End If
 Else
 w_n% = 1
 Call set_new_item(False)
 End If
End If
Call m_display_string.item(w_n%).m_conclusion_or_condition(ty)
     last_input_wenti_type = 10
End Sub

Public Function currentY_position(cond_no%) As Integer
Dim i%, k%
For i% = 1 To cond_no% - 1
k% = m_display_string.item(i%).m_string_no + k%
Next i%
If cond_no% > 0 Then
k% = k% + 1 '最后一条可能是多行,只记第一行
End If
If cond_no% > mvarm_last_conclusion And mvarm_last_conclusion > 0 Then
currentY_position = m_row_space * (k% + 2) + m_pcoordy0
ElseIf cond_no% > mvarm_last_condition And mvarm_last_condition > 0 Then
currentY_position = m_row_space * (k% + 1) + m_pcoordy0
Else
If mvarm_last_condition = 0 And mvarm_last_conclusion = 0 Then
currentY_position = m_row_space * (k% + m_mody_v_position0) + m_pcoordy0
ElseIf cond_no% <= mvarm_last_condition And mvarm_last_condition > 0 Then
currentY_position = m_row_space * k% + m_pcoordy0
ElseIf cond_no% <= mvarm_last_conclusion And mvarm_last_conclusion > 0 Then
currentY_position = m_row_space * (k% + 1) + m_pcoordy0
Else
currentY_position = m_row_space * k% + m_pcoordy0
End If
End If
End Function
Private Function sum_of_row(cond_no2%, cond_no1%) As Integer
Dim i%
For i% = 1 + cond_no1% To cond_no2%
sum_of_row = m_display_string.item(i%).m_string_no + sum_of_row
Next i%
End Function
Public Function currentX_position() As Integer
currentX_position = -m_init_h_position + m_pcoordx0
End Function

Public Sub display_icon(display_or_delete As Byte)
Dim i%
For i% = 1 To m_display_string.Count
 Call m_display_string.item(i%).display_icon(display_or_delete)
Next i%
End Sub
Private Sub Class_Terminate()
  'Set M_InputEvent1 = Nothing
  Set m_display_string = Nothing
End Sub
Public Sub Event_IconDisplayChange(w_n As Integer, ty As Boolean)
 Call set_icon_display(ty)
End Sub

Public Sub Input_Complete(w_n As Integer, string_no As Integer)
    Call control_menu(True) '恢复激活菜单

'             Call set_initial_condition(w_n%, 0, True)
End Sub
Public Function delete_wenti(ty As Integer, new_point%, ele1%, ele2%, inter_point_ty As Integer) As Boolean 'm_no%=1 ,-61
'ty'new_point% 输入新点
If m_display_string.Count > 1 Then
If ty = 1 Then
   If m_display_string.item(m_display_string.Count - 1).is_special_input(1, new_point%, ele1%, ele2%, _
         inter_point_ty) Then
      m_display_string.Remove (m_display_string.Count - 1)
       delete_wenti = True
   End If
ElseIf ty = 7 Then
   If m_display_string.item(m_display_string.Count - 1).is_special_input(7, new_point%, _
         ele1%, ele2%, inter_point_ty) Or _
        m_display_string.item(m_display_string.Count - 1).is_special_input(-61, new_point%, _
         ele1%, ele2%, inter_point_ty) Then
      m_display_string.Remove (m_display_string.Count - 1)
       delete_wenti = True
   End If
ElseIf ty = 9 Then
   If m_display_string.item(m_display_string.Count - 1).is_special_input(9, new_point%, _
        ele1%, ele2%, inter_point_ty) Then
      m_display_string.Remove (m_display_string.Count - 1)
       delete_wenti = True
   End If
ElseIf ty = 11 Then
   If m_display_string.item(m_display_string.Count - 1).is_special_input(11, new_point%, _
         ele1%, ele2%, inter_point_ty) Or _
        m_display_string.item(m_display_string.Count - 1).is_special_input(-63, new_point%, _
         ele1%, ele2%, inter_point_ty) Then
      m_display_string.Remove (m_display_string.Count - 1)
       delete_wenti = True
   End If
End If
End If
End Function

Private Function from_wenti_no_to_index(wenti_no%) As Integer
from_wenti_no_to_index = from_wenti_no_to_index0(wenti_no%, 1)
If from_wenti_no_to_index = 0 Then
   Dim m_dis_string As New display_string
   'If m_display_string.Count Mod 10 = 0 Then
   'ReDim Preserve display_string_s_position(m_display_string.Count + 10) As Integer
   'ReDim Preserve display_string_e_position(m_display_string.Count + 10) As Integer
   'End If
   m_display_string.Add m_dis_string
   'display_string_s_position(m_display_string.Count) = display_string_e_position(m_display_string.Count - 1) + 1
   'display_string_e_position(m_display_string.Count) = display_string_s_position(m_display_string.Count) + 1
   from_wenti_no_to_index = m_display_string.Count
   If wenti_no% = 0 Then
    wenti_no% = m_display_string.Count
   End If
   Call m_display_string.item(m_display_string.Count).set_object(m_object, m_init_h_position, m_init_v_position)
   Call m_display_string.item(m_display_string.Count).set_wenti_no(wenti_no%)
ElseIf wenti_no% = 0 Then
   wenti_no% = from_wenti_no_to_index
   Call m_display_string.item(m_display_string.Count).set_wenti_no(wenti_no%)
End If
End Function
Private Function from_wenti_no_to_index0(ByVal wenti_no%, ty As Byte) As Integer
'ty=0 继续输入,如果wenti_no%=0 nam那么wenti_no%=mvarm_chose_wenti_no,当前操作的语句
'ty=1 新建
Dim i%
If wenti_no% = 0 Then
   If ty = 0 Then
    from_wenti_no_to_index0 = mvarm_chose_wenti_no
   Else
    Exit Function
   End If
Else
  If ty = 0 Then
   For i% = 1 To m_display_string.Count
    If m_display_string.item(i%).wenti_no = wenti_no% Then
     from_wenti_no_to_index0 = i% 'wenti_no对应的序号
      Exit Function
    End If
   Next i%
  End If
End If
End Function
Public Sub set_m_depend_no(wenti_no%)
Call m_display_string.item( _
        from_wenti_no_to_index0(wenti_no%, 0)).set_m_depend_no
End Sub
Public Function m_point_no(ByVal wenti_no%, i%) As Integer
     m_point_no = m_display_string( _
                  from_wenti_no_to_index0(wenti_no%, 0)).m_point_no(i%)

End Function
Public Function m_condition(ByVal no%, i%) As String
     m_condition = m_display_string( _
                  from_wenti_no_to_index0(no%, 0)).m_condition(i%)

End Function
Public Function m_para(ByVal no%, i%) As String
     m_para = m_display_string( _
                  from_wenti_no_to_index0(no%, 0)).para(i%)

End Function

Public Function m_condition_no(ByVal no%) As Integer
     m_condition_no = m_display_string( _
                  from_wenti_no_to_index0(no%, 0)).m_condition_no

End Function

Public Sub set_m_theorem_no(wenti_no%, theorem_no%)
        Call m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).set_m_theorem_no(theorem_no%)
End Sub
Public Sub set_m_condition_data(wenti_no%, ty As Integer, n%)
        Call m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).set_m_condition_data(ty, n%)
End Sub
Public Function m_string(wenti_no%) As String
        m_string = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).m_string

End Function

Public Function m_no(wenti_no%) As Integer
        m_no = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).m_no

End Function

Public Function m_no_(wenti_no%) As Integer
        m_no_ = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).m_no_

End Function
Public Function is_wenti_complete(wenti_no%) As Integer
       is_wenti_complete = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).is_wenti_complete
End Function
Public Function is_last_wenti_complete() As Integer
       If m_display_string.Count > 0 Then
        is_last_wenti_complete = m_display_string.item( _
           m_display_string.Count).is_wenti_complete
       End If
End Function

Public Function m_condition_data_ty(wenti_no%) As Byte
       m_condition_data_ty = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).m_condition_ty

End Function
Public Function m_condition_data_no(wenti_no%) As Integer
       m_condition_data_no = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).m_condition_no

End Function

Public Sub Get_wenti(wenti_no%)
  'If m_display_string.Count > wenti_no% Then
        Call m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).Get_wenti_cond
  'End If
End Sub
Public Function m_depend_no(wenti_no%) As Integer
     m_depend_no = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).m_depend_no

End Function
Public Sub Remove_wenti(wenti_no%)
Dim is_remove As Integer
Dim i%
For i% = 1 To m_display_string.Count
       If m_display_string.item(i% - is_remove).Remove_wenti(wenti_no%) Then
           is_remove = 1
       End If
Next i%
End Sub
Public Function save_wenti_string(wenti_no%) As String
        save_wenti_string = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).set_string
End Function
Public Function m_wenti_theorem_no(wenti_no%) As Integer
        m_wenti_theorem_no = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).m_theorem_no

End Function

Public Function m_condition_ty(wenti_no%) As Byte
        m_condition_ty = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).m_condition_ty

End Function
Public Function m_ty(wenti_no%) As Byte
        m_ty = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).m_ty
End Function
Public Function m_inner_lin(wenti_no%, l%) As Integer
        m_inner_lin = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).m_inner_lin(l%)
End Function
Public Function m_inner_poi(wenti_no%, i%) As Integer
        m_inner_poi = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).m_inner_poi(i%)
End Function
Public Function m_inner_circ(wenti_no%, c%) As Integer
        m_inner_circ = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).m_inner_circ(c%)
End Function


Public Function m_inner_point_type(wenti_no%) As Integer
        m_inner_point_type = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).m_inner_point_type
End Function

Public Sub set_data_condition(wenti_no%, ty As Byte, no%)
        Call m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).set_condition(ty, no%)
End Sub

Public Sub set_m_string(wenti_no%, ByVal brace_mark$, ByVal s1$, ByVal S2$, ByVal s3$, _
                        ByVal note_str$, ByVal theorem_no%, ByVal concl_or_cond%, set_hot_point As Byte) '输入
        Call m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)). _
             set_m_string(brace_mark$, s1$, S2$, s3$, m_total_display_string_no, _
                           note_str$, theorem_no%, concl_or_cond%, set_hot_point)       '输入
End Sub

Public Sub set_m_ty(wenti_no%, ByVal ty As Byte)
        Call m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).set_m_ty(ty)
End Sub

Public Sub Remove_m_string(wenti_no%)
       Call m_display_string.Remove(wenti_no%)
End Sub

Public Sub display_m_string_to_ob(wenti_no%, string_no%, ob As Object, ByVal color%, ByVal display_or_delete As Byte, _
                ByVal m_position%, ByVal set_hot_point As Byte, is_select As Byte, ty As Byte)
        Call m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)). _
             display_m_string_to_ob(string_no%, ob, color%, display_or_delete, _
                 m_position%, set_hot_point, is_select, ty)
End Sub

Public Function m_inpcond(wenti_no%) As String
        m_inpcond = m_display_string.item( _
           from_wenti_no_to_index0(wenti_no%, 0)).m_inpcond

End Function
Public Sub prove_complete()
Dim tn%
tn% = from_wenti_no_to_index0(0, 0)
        Call m_display_string.item(tn%).prove_complete
End Sub

Private Sub set_display_string_position()

End Sub

Public Sub re_display()
Dim i%
Wenti_form.Picture1.Cls
m_total_display_string_no = 0
For i% = 1 To Me.m_last_input_wenti_no
'/////
Next i%
End Sub

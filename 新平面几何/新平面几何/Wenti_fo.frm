VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Wenti_form 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "问　　题"
   ClientHeight    =   10065
   ClientLeft      =   5055
   ClientTop       =   345
   ClientWidth     =   6555
   ControlBox      =   0   'False
   FillColor       =   &H00FFFF80&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   3  'I-Beam
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   671
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   Visible         =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   10610
      _Version        =   393216
      MousePointer    =   1
      Tab             =   1
      TabHeight       =   529
      BackColor       =   16777215
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "传统证明"
      TabPicture(0)   =   "Wenti_fo.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "逆向推理"
      TabPicture(1)   =   "Wenti_fo.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "测量记录"
      TabPicture(2)   =   "Wenti_fo.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture3"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   5652
         Left            =   -75000
         ScaleHeight     =   373
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   333
         TabIndex        =   5
         Top             =   360
         Width           =   5055
         Begin VB.HScrollBar HScroll2 
            Height          =   225
            Left            =   360
            Max             =   200
            TabIndex        =   6
            Top             =   960
            Visible         =   0   'False
            Width           =   1875
         End
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C00000&
         ForeColor       =   &H00C00000&
         Height          =   5655
         Left            =   0
         MousePointer    =   1  'Arrow
         ScaleHeight     =   373
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   325
         TabIndex        =   3
         Top             =   360
         Width           =   4935
         Begin VB.ListBox List1 
            ForeColor       =   &H00FF0000&
            Height          =   1.49820e5
            Left            =   0
            TabIndex        =   10
            Top             =   360
            Width           =   4812
         End
         Begin ComctlLib.TreeView TreeView1 
            Height          =   3972
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   2400
            Width           =   4812
            _ExtentX        =   8493
            _ExtentY        =   7011
            _Version        =   327682
            LineStyle       =   1
            Style           =   6
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawMode        =   7  'Invert
         FillColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15655
         Left            =   -74880
         ScaleHeight     =   1040
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   317
         TabIndex        =   1
         Top             =   360
         Width           =   4815
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000009&
            BorderStyle     =   0  'None
            Height          =   1410
            Left            =   1440
            Picture         =   "Wenti_fo.frx":0054
            ScaleHeight     =   94
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   77
            TabIndex        =   9
            Top             =   2760
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   3375
            Left            =   4440
            Max             =   100
            TabIndex        =   8
            Top             =   0
            Width           =   255
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   120
            Max             =   50
            TabIndex        =   7
            Top             =   5160
            Width           =   3495
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000018&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1440
            TabIndex        =   2
            Top             =   1920
            Visible         =   0   'False
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "Wenti_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim old_VS_value%
Dim time_no As Integer
Dim top_of_treeview%
Dim number As Integer
Dim mouse_y  As Single
'Dim cond_no() As condition_no_type


Private Sub Form_Activate()
If arrange_window_type = 2 Then
 'arrange_window_type = 2
Wenti_form.left = 100
Wenti_form.top = 320
Wenti_form.width = Screen.width - 280
Wenti_form.Height = Screen.Height - 1550 + int_w_y
Draw_form.width = Screen.width - 280
Draw_form.left = 0
Draw_form.top = 0
Draw_form.Height = Screen.Height - 1550 + int_w_y
 Draw_form.Picture1.Height = Draw_form.ScaleHeight
  Draw_form.Picture1.width = Draw_form.ScaleWidth
 Draw_form.Refresh
   Wenti_form.Refresh
    Wenti_form.SetFocus
End If
End Sub
'输入可读字符
' 键码
Private Sub Form_Load()
input_condition_no = 0
'Text2.Top = 0
'Text2.Height = 98
Picture1.top = 360
Picture1.left = 0
Picture1.CurrentX = 0
Picture1.CurrentY = 0
time_no = 0
modify_statue = no_modify
display_theorem = False
'record0.is_proved = 1
TreeView1(0).Height = Picture2.Height - TreeView1(0).top
TreeView1(0).width = Picture2.width - TreeView1(0).left
List1.top = 25
'**************************
SSTab1.Tab = 2
SSTab1.Caption = LoadResString_(4235, "")
SSTab1.Tab = 1
SSTab1.Caption = LoadResString_(4230, "")
SSTab1.Tab = 0
SSTab1.Caption = LoadResString_(1905, "")
SSTab1_name_type = 0
Me.Caption = LoadResString_(4240, "")
End Sub

Private Sub Form_Resize()
If arrange_window_type = 0 Then
SSTab1.width = Wenti_form.ScaleWidth
SSTab1.Height = Wenti_form.ScaleHeight
Picture1.Height = Wenti_form.Height
Picture1.width = Wenti_form.width
Picture2.Height = Wenti_form.Height
Picture2.width = Wenti_form.width
Picture3.Height = Wenti_form.Height
Picture3.width = Wenti_form.width
List1.width = Wenti_form.ScaleWidth - 10
''Draw_form.Width = Wenti_form.Left - 5
 If Screen.width > Draw_form.width Then
 Wenti_form.width = Screen.width - Draw_form.width
 End If
  Wenti_form.top = 0
   Wenti_form.Height = Screen.Height - 1350 + int_w_y
If Picture1.width + Picture1.left < ScaleWidth Then
Picture1.width = ScaleWidth - Picture1.left
End If
If Picture1.Height + Picture1.top < ScaleHeight - 20 Then
Picture1.Height = ScaleHeight - Picture1.top - 20
End If
ElseIf arrange_window_type = 1 Then
Draw_form.top = 0
Draw_form.left = 0
Wenti_form.left = 0
Draw_form.width = Screen.width - 150
Wenti_form.width = Screen.width - 150
If Wenti_form.top - 5 > 0 Then
Draw_form.Height = Wenti_form.top - 5
End If
If Screen.Height - 1610 + int_w_y - Wenti_form.top > 0 Then
Wenti_form.Height = Screen.Height - 1350 + int_w_y - Wenti_form.top
End If
End If
VScroll1.left = ScaleWidth - 20
VScroll1.Height = ScaleHeight - 42
HScroll1.top = VScroll1.Height ' - Picture1.Top - 16
HScroll1.width = ScaleWidth - 20
 Draw_form.Picture1.Height = Draw_form.ScaleHeight
  Draw_form.Picture1.width = Draw_form.ScaleWidth
  
  Wenti_form.HScroll1.top = Wenti_form.ScaleHeight - 68
Wenti_form.VScroll1.Height = Wenti_form.ScaleHeight - 68
  
End Sub
Private Sub HScroll1_Change()
Call C_display_wenti.m_change_display(HScroll1.value, 0)
End Sub

Private Sub HScroll2_Change()
Dim m%, i%
Dim n As Byte
If Ratio_for_measure.is_fixed_ratio Then
 Exit Sub
Else
If HScroll2.value > 1 And HScroll2.value < 480 Then
m% = HScroll2.value
If is_set_Hscroll2_data Then
 Call Wenti_form.Picture3.Cls
  Ratio_for_measure.Ratio_for_measure = m%
   Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
   Call change_ratio_for_measure
'    For i% = 1 To Ratio_for_measure.sons.last_son
'     If Ratio_for_measure.sons.son(i%).ty = line_ Then
'     ElseIf Ratio_for_measure.sons.son(i%).ty = circle_ Then
'      Call change_m_circle(Ratio_for_measure.sons.son(i%).no, depend_condition(Ratio_for_measure_, 0))
'     ElseIf Ratio_for_measure.sons.son(i%).ty = wenti_cond_ Then
'     End If
' Next i%
    Call measur_again
    ' Call C_display_picture.Backup_picture
'Call change_picture(0, 0)
'Call draw_again0(Draw_form, 0)
End If
End If
End If
End Sub

Private Sub list1_Click()
 Call create_treeview(inform_data_base(List1.ListIndex).ty, _
             inform_data_base(List1.ListIndex).no, TreeView1(0), _
              List1.Height + 30)
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
 mouse_y = Y
End If
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
List1.top = List1.top + mouse_y - Y
mouse_y = Y
End If
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim k%, i%
Dim char As String
'On Error GoTo picture1_keydown_error
If KeyCode = 13 Then
 next_step_of_profe = True
 Call C_display_wenti.m_input_char(Wenti_form.Picture1, Chr(13))
 Exit Sub
ElseIf KeyCode = 38 Then 'up
'If VScroll1.value > 0 Then
'If C_display_wenti.m_change_display(Picture1, 0, VScroll1.value - 1) Then
' VScroll1.value = VScroll1.value - 1
'End If
'End If
' Exit Sub
ElseIf KeyCode = 40 Then 'dwon
'If C_display_wenti.m_change_display(Picture1, 0, VScroll1.value + 1) Then
' VScroll1.value = VScroll1.value + 1
'End If
' Exit Sub
ElseIf KeyCode = 39 Then 'right
'If HScroll1.value > 0 Then
'If C_display_wenti.m_change_display(Picture1, 20 * (HScroll1.value - 1), 0) Then
' HScroll1.value = HScroll1.value - 1
'End If
'End If
' Exit Sub
ElseIf KeyCode = 37 Then 'left
'If C_display_wenti.m_change_display(Picture1, 20 * (HScroll1.value + 1), 0) Then
' HScroll1.value = HScroll1.value + 1
'End If
' Exit Sub
ElseIf KeyCode = 8 Then 'back
 Call C_display_wenti.m_input_char(Wenti_form.Picture1, "backspace")
  Exit Sub
End If
If run_type < 1 Or regist_data.set_password_Checked = False Then
code = KeyCode '**
char = Chr(code)
'Call C_display_wenti.m_input_char(Wenti_form.Picture1, char)
Exit Sub
Else
 last_conditions.last_cond(1).pass_word_for_teacher = _
      Mid$(last_conditions.last_cond(1).pass_word_for_teacher, 2, 4) + LCase(Chr(KeyCode))
   k% = InStr(1, last_conditions.last_cond(1).pass_word_for_teacher, "*", 0)
    MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1240, "")
 For i% = 1 To 5 - k%
  MDIForm1.StatusBar1.Panels(1).text = MDIForm1.StatusBar1.Panels(1).text + "*"
 Next i%
  If k% = 0 Then
   If Mid$(last_conditions.last_cond(1).pass_word_for_teacher, 1, 5) = _
        protect_data.pass_word_for_teacher Then
          Call C_display_wenti.display_result(Wenti_form.Picture1, True)
   Else
    MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1245, "")
    last_conditions.last_cond(1).pass_word_for_teacher = "0000*"
   End If
  End If
 End If
'Picture1.SetFocus
picture1_keydown_error:

End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
Dim ch As String
ch = Chr(KeyAscii) '输入的字符
If ch$ = "@" Then
 ch = "`"
ElseIf ch$ = "#" Then
 ch = "\"
ElseIf Asc(ch) = 13 Then
 ch = "~"
ElseIf ch <= "9" And ch >= "0" Then
 ch = ch
ElseIf (ch < "a" Or ch > "Z") And ch <> "+" And ch <> "-" And ch <> "*" And ch <> "/" Then
 Exit Sub
End If
 Call C_display_wenti.m_input_char(Wenti_form.Picture1, ch)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, _
     X As Single, Y As Single)
Dim i%
'Call C_display_wenti.get_MouseDown(Button, Int(X), Int(Y))
If run_type < 1 Then
'设置小热键菜单
 If Button = 2 And X > 20 Then
  Picture4.top = Y
  Picture4.left = X
   If Picture4.visible = False Then
    Picture4.visible = True
   End If
  Exit Sub
 Else
  Picture4.visible = False
 End If
End If
'End If
'选语句
If protect_data.pass_word_for_teacher = "00000" Or _
       InStr(1, last_conditions.last_cond(1).pass_word_for_teacher, "*", 0) = 0 Then
'Call C_display_wenti.click_display_string(Picture1, Int(X), Int(Y), modify_input_statue)
If C_display_wenti.get_MouseDown(Button, CInt(X), CInt(Y), chose_w_no%) = False Then
 Exit Sub
End If
If modify_input_statue = True Then
Call back_input_to(chose_w_no%)
 Exit Sub
End If
 If chose_w_no% <> select_wenti_no% Then
 If chose_w_no% >= 0 Then
'  If select_wenti_no% > 0 Then
'  Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, _
          1, select_wenti_no%, 0, 0, 0, 2)
'  End If
  select_wenti_no% = chose_w_no%
  'Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, _
         1, chose_w_no%, 0, 0, 1, 1)
  Call draw_inform(chose_w_no%, C_display_wenti.m_condition_ty(chose_w_no%), _
         C_display_wenti.m_condition_no(chose_w_no%))
 End If
 End If
'End If
'**********************
If run_statue = 4 Then '12.10
 If C_display_wenti.get_MouseDown(Button, CInt(X), CInt(Y), chose_w_no%) Then
  If wenti_cond_no_reduce(chose_w_no%) = False Then
      wenti_cond_no_reduce(chose_w_no%) = True
    Call init_condition
    For i% = 1 To C_display_wenti.m_last_input_wenti_no
     If wenti_cond_no_reduce(i%) = False Then
      Call draw_picture(i%, 0, True)
     Else
      Call draw_picture(i%, 255, True)
     End If
    Next i%
    MDIForm1.Toolbar1.Buttons(19).visible = True
    MDIForm1.Toolbar1.Buttons(17).visible = False
    MDIForm1.Toolbar1.Buttons(18).visible = False
  Else
 ' Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, 0, chose_w_no%, chose_w_no%, _
            1, 0, 0)
  'Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, 1, chose_w_no%, chose_w_no%, _
            1, 0, 0)
   wenti_cond_no_reduce(chose_w_no%) = False
    Call draw_picture(chose_w_no%, 0, True)
    MDIForm1.Toolbar1.Buttons(19).visible = True
    MDIForm1.Toolbar1.Buttons(17).visible = False
    MDIForm1.Toolbar1.Buttons(18).visible = False
   End If
 End If
Else
'*******************************

If C_display_wenti.m_last_input_wenti_no > 0 Then
If C_display_wenti.get_MouseDown(Button, CInt(X), CInt(Y), 0) And Button = 2 And _
     modify_wenti_no = C_display_wenti.m_last_input_wenti_no And _
      draw_wenti_no <= C_display_wenti.m_last_input_wenti_no Then
       Call C_display_wenti.Get_wenti(modify_wenti_no)
 temp_wenti_cond(modify_wenti_no) = wenti_cond0.data 'C_display_wenti.m_display_string.item(modify_wenti_no)
  '保存修改语句,最后一句
  If modify_condition_no = -1 Then
   event_statue = wait_for_modify_sentence
    'icon_char = str(modify_wenti_no + 1)
    ' icon_x = -2
    '  icon_y = 20 * (modify_wenti_no + 1)
    '          time_no = 0
      MDIForm1.Timer1.interval = 500
      MDIForm1.Timer1.Enabled = True
 Else
  event_statue = wait_for_modify_char
        time_no = 0
    'icon_x = modify_icon_x
    'icon_y = modify_icon_y
    'icon_position = modify_icon_fontsize
    MDIForm1.Timer1.Enabled = True
 End If
Exit Sub
End If
End If
End If
End If
End Sub
Public Sub display_inform_in_treeview(ty As Byte, n%, is_root As Boolean)
Dim i%
Dim nod As Node
Dim temp_record As record_type
Dim ind As String
End Sub

Private Sub Picture1_Resize()
VScroll1.top = 0 'Picture1.Top
VScroll1.left = Wenti_form.ScaleWidth - 16
VScroll1.Height = Wenti_form.ScaleHeight - 16
HScroll1.top = Wenti_form.ScaleHeight - 16
HScroll1.width = Wenti_form.ScaleWidth
HScroll1.left = -Picture1.left
End Sub

Private Sub picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim date_type As Byte
Dim date_no%, no%, i%, j%, con_no%
Dim nod As Node
Dim ind As String
Dim temp_record As total_record_type
If Button = 1 Then
If X > 0 And X < 200 Then
If Y > 21 And Y < 40 Then
 no% = 0
ElseIf Y > 41 And Y < 60 Then
 no% = 1
ElseIf Y > 61 And Y < 80 Then
 no% = 2
ElseIf Y > 81 And Y < 100 Then
 no% = 3
Else
 Exit Sub
End If
End If
If last_conclusion > no% Then
 If conclusion_data(no%).ty > 0 And conclusion_data(no%).no(0) <> 0 Then
  last_dis_con_gs(dis_gs_no%) = 0
  Last_hotpoint_of_theorem1 = 0
   Erase Hotpoint_of_theorem1
    last_node_index = 0
  '设置TreeView的几何大小
 '**********************************************************
 Call create_treeview(conclusion_data(no%).ty, conclusion_data(no%).no(no%), TreeView1(0), top_of_treeview% + 20)
 'If conclusion_data(no%) = general_string_ Then
 ' If con_general_string(no%).data(0).value = "" Then
 '         j% = conclusion_no(no%, 0)
 '             last_dis_con_gs(no%) = 1
 '              con_g_s(no%, last_dis_con_gs(no%)) = j%
 '               j% = general_string(j%).data(0).record.data0.condition_data.condition( _
 '                    general_string(j%).data(0).record.data0.condition_data.condition_no).no
 '           Do While general_string(j%).data(0).value = ""
 '            last_dis_con_gs(no%) = 1 + last_dis_con_gs(no%)
 '             con_g_s(no%, last_dis_con_gs(no%)) = j%
 '             If general_string(j%).data(0).record.data0.condition_data.condition_no < 9 Then
 '                  If general_string(j%).data(0).record.data0.condition_data.condition( _
 '                    general_string(j%).data(0).record.data0.condition_data.condition_no).ty = general_string_ Then
 '                     j% = general_string(j%).data(0).record.data0.condition_data.condition( _
 '                       general_string(j%).data(0).record.data0.condition_data.condition_no).no
 '                  Else
 '                   last_dis_gs(no%) = last_dis_con_gs(no%)
 '                    GoTo mousedown_mark1
 '                  End If
 '              Else
 '               last_dis_gs(no%) = last_dis_con_gs(no%)
 '                dis_gs_no% = no%
 '                GoTo mousedown_mark1
 '              End If
 '          Loop
 '          last_dis_gs(dis_gs_no%) = last_dis_con_gs(dis_gs_no%)
 '          GoTo mousedown_mark1
 '  Else
 '  date_type = conclusion_data(no%)
 '  date_no% = conclusion_no(no%, 0)
 ' End If
 'Else
 '  date_type = conclusion_data(no%)
 '  date_no% = conclusion_no(no%, 0)
 'End If
'mousedown_mark1:
'If last_dis_con_gs(dis_gs_no%) > 0 Then
'    date_type = general_string_
'     date_no% = con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))
'End If
'**********************************************************************************
'TreeView1(0).visible = True
'TreeView1(0).Nodes.Clear
'last_node_index = 1
'ind = "node" & CStr(last_node_index)
'Set nod = TreeView1(0).Nodes.Add(, , ind, _
'            set_display_inform(set_display_string0(date_type, date_no%, 0, False, False, 0, 1, 0), date_type))
'ReDim Preserve cond_no(nod.index) As condition_no_type
'cond_no(nod.index).ty = date_type
'cond_no(nod.index).no = date_no%
'If date_type = general_string_ Then
' If general_string(date_no%).data(0).value = "" Then
'     last_dis_gs(dis_gs_no%) = last_dis_gs(dis_gs_no%) - 1
'      If last_dis_gs(dis_gs_no%) > 0 Then
'       Call add_node(TreeView1(0), ind, general_string_, con_g_s(dis_gs_no%, last_dis_gs(dis_gs_no%)))
'       If general_string(con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))).data(0).record.data0.condition_data.condition_no > 0 And _
'              general_string(con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))).data(0).record.data0.condition_data.condition_no < 9 Then
'        For i% = 1 To general_string(con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))).data(0).record.data0.condition_data.condition_no - 1
'         Call record_no(general_string(con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))).data(0).record.data0.condition_data.condition(j%).ty, _
'               general_string(con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))).data(0).record.data0.condition_data.condition(j%).no, _
'                   temp_record, True, 0, 0)
'         Call add_node(TreeView1(0), ind, _
'           general_string(con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))).data(0).record.data0.condition_data.condition(i%).ty, _
'              general_string(con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))).data(0).record.data0.condition_data.condition(i%).no)
'        Next i%
'         con_no% = general_string(con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))).data(0).record.data0.condition_data.condition_no
'        If general_string(con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))).data(0).record.data0.condition_data.condition(con_no).ty = general_string_ Then
'         If general_string(general_string(con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))).data(0).record.data0.condition_data.condition(con_no%).no). _
'              data(0).value = "" Then
'
'         Else
'         Call add_node(TreeView1(0), ind, _
'           general_string(con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))).data(0).record.data0.condition_data.condition(con_no%).ty, _
'              general_string(con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))).data(0).record.data0.condition_data.condition(con_no%).no)
'         End If
'        Else
'         Call add_node(TreeView1(0), ind, _
'           general_string(con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))).data(0).record.data0.condition_data.condition(con_no%).ty, _
'              general_string(con_g_s(dis_gs_no%, last_dis_con_gs(dis_gs_no%))).data(0).record.data0.condition_data.condition(con_no%).no)
'        End If
'       End If
'      End If
' Else
'  Call record_no(date_type, date_no%, temp_record, True, 0, 0)
'  If temp_record.record_data.data0.condition_data.condition_no > 0 And temp_record.record_data.data0.condition_data.condition_no < 9 Then
'   For i% = 1 To temp_record.record_data.data0.condition_data.condition_no
'    Call add_node(TreeView1(0), ind, _
'       temp_record.record_data.data0.condition_data.condition(i%).ty, temp_record.record_data.data0.condition_data.condition(i%).no)
'   Next i%
'  End If
' End If
'Else
' Call record_no(date_type, date_no%, temp_record, True, 0, 0)
'  If temp_record.record_data.data0.condition_data.condition_no > 0 And temp_record.record_data.data0.condition_data.condition_no < 9 Then
'   For i% = 1 To temp_record.record_data.data0.condition_data.condition_no
'    Call add_node(TreeView1(0), ind, _
'       temp_record.record_data.data0.condition_data.condition(i%).ty, temp_record.record_data.data0.condition_data.condition(i%).no)
'   Next i%
'  End If
'End If
 End If
 End If
End If
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ch As String
If inp = 50 Or inp = -49 Or inp = 38 Or inp = 22 Then
If Button = 1 Then
 If condition_no > 0 And Y < 63 Then
 If C_display_wenti.m_condition(modify_wenti_no, condition_no - 1) = "$" Or _
     C_display_wenti.m_condition(modify_wenti_no, condition_no - 1) = "&" Or _
      C_display_wenti.m_condition(modify_wenti_no, condition_no - 1) = "`" Or _
       C_display_wenti.m_condition(modify_wenti_no, condition_no - 1) = "\" Then
         Call Picture1.SetFocus
    Exit Sub
 End If
 End If
 If X > 0 And X < 38 Then
 If Y > 19 And Y < 30 Then
 ch = "$" 'sin
 ElseIf Y > 30 And Y < 41 Then
 ch = "&" 'cos
 ElseIf Y > 41 And Y < 53 Then
 ch = "`" 'tan
 ElseIf Y > 53 And Y < 63 Then
 ch = "\" 'ctan
 ElseIf Y > 63 And Y < 73 Then
 ch = Chr(12) ' "," 'angle
 ElseIf Y > 73 And Y < 85 Then
 ch = "'" 'sqr_root LoadResString_(772)
 Else
 ch = ""
 End If
 Call C_display_wenti.m_input_char(Wenti_form.Picture1, ch)
 End If
End If
End If
         Call Picture1.SetFocus
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
Dim i%, j%
Dim y_position%
Dim t_SSTab1_Tab As Integer
t_SSTab1_Tab = SSTab1.Tab
If SSTab1_name_type = 1 And t_SSTab1_Tab <> 1 Then '0
   SSTab1.Tab = 1
   SSTab1.Caption = LoadResString_(4230, "")
   SSTab1_name_type = 0
   List1.visible = False
End If
'SSTab1.Tab = t_SSTab1_Tab
If regist_data.run_type = 0 Then
'***********************************************
If SSTab1.Tab = 0 Then '显示传统证明
  Picture1.visible = True
'*********************************************
ElseIf SSTab1.Tab = 1 Then '显示推理树
  Picture2.Cls
   If List1.visible Then '
     Call SetTextColor(Picture2.hdc, QBColor(9))
     Picture2.Print database_name
     top_of_treeview% = List1.Height
   Else
     Call SetTextColor(Picture2.hdc, QBColor(9))
     Picture2.Print LoadResString_(1320, "")
      If run_statue >= 6 Then '12.10
        For i% = 1 To C_display_wenti.m_last_input_wenti_no
         If C_display_wenti.m_no(i%) > 22 Then
          y_position% = y_position% + 20
          Call C_display_wenti.display_m_string_to_ob(i%, 0, Wenti_form.Picture2, _
           conclusion_color, 1, y_position%, 0, 0, 1)
         End If
     Next i%
   End If
     top_of_treeview% = Picture2.CurrentY
  End If
'************************************************************
ElseIf SSTab1.Tab = 2 Then '测量
   Picture3.Cls
   If Ratio_for_measure.Ratio_for_measure > 0 And is_set_Hscroll2_data = False Then '有关于长度(或面积)的条件
     Wenti_form.HScroll2.min = Ratio_for_measure.ratio_for_measure0 / 2
     Wenti_form.HScroll2.max = Ratio_for_measure.ratio_for_measure0 * 2
     Wenti_form.HScroll2.value = Ratio_for_measure.ratio_for_measure0
     is_set_Hscroll2_data = True
   End If
    Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
    HScroll2.visible = True
    Call measur_again
   End If
End If
End Sub
Private Sub TreeView1_BeforeLabelEdit(index As Integer, Cancel As Integer)
 Me.Cls
End Sub
Private Sub TreeView1_NodeClick(index As Integer, ByVal Node As ComctlLib.Node)
Dim temp_record As total_record_type
Dim n%
If Node.index <= last_node_index Then
Call record_no(cond_no(Node.index).ty, cond_no(Node.index).no, temp_record, True, 0, 0)
n% = temp_record.record_data.data0.theorem_no
Call draw_inform(0, cond_no(Node.index).ty, cond_no(Node.index).no)
 If n% > 0 Then
MDIForm1.StatusBar1.Panels(1).text = th_chose(n%).text
 End If
End If
End Sub
Private Sub VScroll1_Change()
Call C_display_wenti.m_change_display(0, VScroll1.value)
End Sub

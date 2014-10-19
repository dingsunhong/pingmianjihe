VERSION 5.00
Begin VB.Form set_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置 "
   ClientHeight    =   2805
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   3360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "信息树显示"
      Height          =   252
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   3855
   End
   Begin VB.CheckBox Check8 
      Caption         =   "向量法"
      Height          =   192
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   3855
   End
   Begin VB.CheckBox Check7 
      Caption         =   "传统证明方法"
      Height          =   252
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   3855
   End
   Begin VB.CheckBox Check6 
      Caption         =   "保留原题不动"
      Height          =   252
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   3855
   End
   Begin VB.CheckBox Check5 
      Caption         =   "显示用法提示小窗口"
      Height          =   252
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   3855
   End
   Begin VB.CheckBox Check4 
      Caption         =   "设置监护密码"
      Height          =   252
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   3855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "投影仪方式"
      Height          =   192
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "工具栏"
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "set_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim old_language As Byte
Private Sub Check1_Click()
If set_form.Check1.value = 0 Then
  int_w_y = 400
  If arrange_window_type = 0 Then
  Draw_form.Height = screeny& - 1350 + int_w_y
   Wenti_form.Height = screeny& - 1350 + int_w_y
  ElseIf arrange_window_type = 1 Then
   Wenti_form.Height = screeny& - 1350 + int_w_y - Wenti_form.top
  End If
      MDIForm1.Toolbar1.visible = False
        regist_data.Tool_bar_visible = False
Else
   int_w_y = 0
    MDIForm1.Toolbar1.visible = True
    regist_data.Tool_bar_visible = True
 If arrange_window_type = 0 Then
Draw_form.Height = screeny& - 1350 + int_w_y
 Wenti_form.Height = screeny& - 1350 + int_w_y
 ElseIf arrange_window_type = 1 Then
  If screeny& - 1610 + int_w_y - Wenti_form.top > 0 Then
  Wenti_form.Height = screeny& - 1350 + int_w_y - Wenti_form.top
  Else
  Draw_form.Height = screeny& - 2100
  Wenti_form.top = screeny& + 5
  Wenti_form.Height = screeny& - 1350 + int_w_y - Wenti_form.top
  End If
   End If
End If
 Draw_form.Picture1.Height = Draw_form.ScaleHeight
  Draw_form.Picture1.width = Draw_form.ScaleWidth
Call set_regist
End Sub

Private Sub Check2_Click()
If set_form.Check2.value = 1 Then
    Call init_project_(True)
Else
    Call init_project_(False)
End If
Call set_regist
End Sub
Private Sub Check3_Click()
  If Check3.value = 1 Then
     Call init_tree_(True)
  Else 'If Check3.value = 1 Then
     'Check3.value = 0
     Call init_tree_(False)
  End If
  Call set_regist
End Sub
Private Sub Check4_Click()
password.Text1.text = "*****"
password.Show
If Check4.value = 1 Then
password.Caption = LoadResString_(4110, "")
 'protect_data.pass_word_for_teacher = Mid$(password.Text1.text, 1, 5)
  'regist_data.set_password_Checked = True
   'regist_data.teacher_password = protect_data.pass_word_for_teacher
Else 'If Me.treeview.Checked Then
  password.Caption = LoadResString_(4115, "")
End If
If Check4.value = 1 And protect_data.pass_word_for_teacher <> "00000" Then
Call C_display_wenti.set_m_pass_word_for_teacher(False)
Else
Call C_display_wenti.set_m_pass_word_for_teacher(True)
End If
Call set_regist
End Sub

Private Sub Check5_Click()
 If Check5.value = 1 Then
 '   Check5.value = 1
     regist_data.mnuTip_Checked = True
 Else
 ' Check5.value = 0
      regist_data.mnuTip_Checked = False
 End If
 Call set_regist
End Sub
Private Sub Check6_Click()
If Check6.value = 1 Then
regist_data.yuantibudong_Checked = True
Call C_display_wenti.set_m_yuantibudong(True)
Else
regist_data.yuantibudong_Checked = False
Call C_display_wenti.set_m_yuantibudong(False)
End If
Call set_regist
End Sub
Private Sub Check7_Click()
 If Check7.value = 0 Then
    Check8.value = 1
    regist_data.run_type = 1
        MDIForm1.Caption = Mdiform1_caption + "-" + LoadResString_(1135, "")
 Else
    regist_data.run_type = 0
     Check8.value = 0
        MDIForm1.Caption = Mdiform1_caption + "-" + LoadResString_(1140, "") ' "版—传统证明方法"
 End If
 Call set_regist
End Sub

Private Sub Check8_Click()
 MDIForm1.Caption = LoadResString_(110, "") & App.Major & "." & App.Minor & "." & App.Revision
 If Check8.value = 0 Then
    Check7.value = 1
    Check7.Enabled = True
    regist_data.run_type = 0
    MDIForm1.Caption = MDIForm1.Caption + "-" + LoadResString_(1135, "")
 Else
     Check7.value = 0
    regist_data.run_type = 1
        MDIForm1.Caption = MDIForm1.Caption + "-" + LoadResString_(1140, "")
 End If
 Call set_regist
End Sub
Private Sub Combo1_Click()
old_language = regist_data.language
'If Combo1.ListIndex = 0 Then
   regist_data.language = Combo1.ListIndex + 1
'ElseIf Combo1.ListIndex = 1 Then
'   regist_data.language = 2
'ElseIf Combo1.ListIndex = 2 Then
'   regist_data.language = 3
'End If
'If old_language <> regist_data.language Then
'    Call MDIForm1.Set_inpcond
'    Call MDIForm1.Set_Mune_Item
'    If regist_data.run_type = 0 Then
'       MDIForm1.Caption = Mdiform1_caption + "-" + LoadResString_(1135, "")
'    ElseIf regist_data.run_type = 1 Then
'       MDIForm1.Caption = MDIForm1.Caption + "-" + LoadResString_(1140, "") ' "版—传统证明方法"
'    End If
'    Call set_caption
'    Draw_form.Caption = LoadResString_(2005, "") + "-" + LoadResString_(1925, "")
'    Wenti_form.Caption = LoadResString_(1960, "") + "-" + LoadResString_(1925, "") + _
'                         LoadResString_(3955, "\\1\\" + LoadResString_(425, ""))
'    Wenti_form.SSTab1.Tab = 2
'    Wenti_form.SSTab1.Caption = LoadResString_(4235, "")
'    Wenti_form.SSTab1.Tab = 1
'    Wenti_form.SSTab1.Caption = LoadResString_(4230, "")
'    Wenti_form.SSTab1.Tab = 0
'    Wenti_form.SSTab1.Caption = LoadResString_(1905, "")
'    Call C_display_wenti.re_display
'    Call C_IO.set_exam_list
'    Call C_IO.add_exam_name_to_list(exam_form.List1)
'End If
End Sub

Private Sub Command1_Click()
If old_language <> regist_data.language Then
    Call MDIForm1.Set_inpcond
    Call MDIForm1.Set_Mune_Item
    If regist_data.run_type = 0 Then
       MDIForm1.Caption = Mdiform1_caption + "-" + LoadResString_(1135, "")
    ElseIf regist_data.run_type = 1 Then
       MDIForm1.Caption = MDIForm1.Caption + "-" + LoadResString_(1140, "") ' "版—传统证明方法"
    End If
    Call set_caption
    Draw_form.Caption = LoadResString_(2005, "") + "-" + LoadResString_(1925, "")
    Wenti_form.Caption = LoadResString_(1960, "") + "-" + LoadResString_(1925, "") + _
                         LoadResString_(3955, "\\1\\" + LoadResString_(425, ""))
    Wenti_form.SSTab1.Tab = 2
    Wenti_form.SSTab1.Caption = LoadResString_(4235, "")
    Wenti_form.SSTab1.Tab = 1
    Wenti_form.SSTab1.Caption = LoadResString_(4230, "")
    Wenti_form.SSTab1.Tab = 0
    Wenti_form.SSTab1.Caption = LoadResString_(1905, "")
    Call C_display_wenti.re_display
    Call C_IO.set_exam_list
    Call C_IO.add_exam_name_to_list(exam_form.List1)
End If
set_form.Hide
End Sub

Private Sub Command2_Click()
  regist_data.language = old_language
  Combo1.ListIndex = regist_data.language - 1
  set_form.Hide
End Sub

Private Sub Form_Load()
Combo1.AddItem LoadResString(806)
Combo1.AddItem LoadResString(807)
Combo1.AddItem LoadResString(808)
Call set_caption
End Sub

Private Sub set_caption()
set_form.Caption = LoadResString_(145, "")
Check1.Caption = LoadResString_(825, "")
Check2.Caption = LoadResString_(1735, "")
Check3.Caption = LoadResString_(1740, "")
Check4.Caption = LoadResString_(1610, "")
Check5.Caption = LoadResString_(1750, "")
Check6.Caption = LoadResString_(1900, "")
Check7.Caption = LoadResString_(1135, "")
Check8.Caption = LoadResString_(1140, "")
'If regist_data.language = 1 Then
   Combo1.text = LoadResString_(805, "")
'ElseIf regist_data.language = 2 Then
'   Combo1.text = LoadResString(807)
'ElseIf regist_data.language = 3 Then
'   Combo1.text = LoadResString(808)
'End If
Me.Command1.Caption = LoadResString_(3940, "")
Me.Command2.Caption = LoadResString_(135, "")
End Sub


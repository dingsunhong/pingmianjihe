VERSION 5.00
Begin VB.Form password 
   Caption         =   "设置监护密码"
   ClientHeight    =   1752
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4188
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1752
   ScaleWidth      =   4188
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Text            =   "00000"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "请输入监护密码:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   612
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   1692
   End
End
Attribute VB_Name = "password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim data_con As String * 60
If set_form.Check4.value = 1 Then
If protect_data.pass_word_for_teacher = "00000" Then
  Me.Text1.text = Mid$(Me.Text1.text, 1, 5)
 If Me.Text1.text <> "00000" And (InStr(1, Me.Text1.text, "*", 0) = 0 Or _
        InStr(1, Me.Text1.text, "*", 0) > 5) Then
 protect_data.pass_word_for_teacher = LCase(Mid$(Me.Text1.text, 1, 5))
   inform.TreeView1.visible = True
  inform_treeview_visible = True
   inform_picture_visible = False
    inform.VScroll1.visible = False
'     MDIForm1.TreeView.Checked = True
   regist_data.set_password_Checked = True
    regist_data.teacher_password = protect_data.pass_word_for_teacher
     set_form.Check4.value = 1
    data_con = set_protect_data_to_string(protect_data)
     Open protect_file(0) For Binary As #4 '打开sound
      Put #4, 103, data_con  '读出保护信息
       Close #4
Call put_filetime(0)
Open protect_file(1) For Binary As #6 '打开dll
Put #6, 1, data_con  '读出保护信息
Close #6
Call put_filetime(1)
Call put_filetime(2)

    Me.Hide
 Else
 MsgBox LoadResString_(1870, ""), , LoadResString_(1865, "")
 End If
Else
 If protect_data.pass_word_for_teacher = LCase(Mid$(Me.Text1.text, 1, 5)) Then
  protect_data.pass_word_for_teacher = "00000"
 inform.TreeView1.visible = False
  inform_treeview_visible = False
   inform_picture_visible = True
    inform.VScroll1.visible = True
 ' MDIForm1.TreeView.Checked = False
   regist_data.set_password_Checked = False
    regist_data.teacher_password = protect_data.pass_word_for_teacher
     set_form.Check4.value = 0
     data_con = set_protect_data_to_string(protect_data)
     Open protect_file(0) For Binary As #4 '打开sound
      Put #4, 103, data_con  '读出保护信息
       Close #4
Call put_filetime(0)
Open protect_file(1) For Binary As #6 '打开dll
Put #6, 1, data_con  '读出保护信息
Close #6
Call put_filetime(1)
'Call put_filetime(2)
    Me.Hide
Else
    MsgBox LoadResString_(1870, ""), , LoadResString_(1875, "")
 End If
End If
Else
   If protect_data.pass_word_for_teacher = Mid$(password.Text1.text, 1, 5) Then
     protect_data.pass_word_for_teacher = "000000"
      regist_data.teacher_password = protect_data.pass_word_for_teacher
       regist_data.set_password_Checked = False
   Else
    set_form.Check4.value = 1
   End If
   Me.Hide
End If
End Sub

Private Sub Command2_Click()
 If set_form.Check4.value = 1 Then
     regist_data.set_password_Checked = False
      set_form.Check4.value = 0
 Else
   set_form.Check4.value = 1
     regist_data.set_password_Checked = True
 End If
  Me.Hide
End Sub

Private Sub Form_Load()
Me.Caption = LoadResString_(1610, "")
Me.Label1.Caption = LoadResString_(4250, "")
Me.Command1.Caption = LoadResString_(3940, "")
Me.Command2.Caption = LoadResString_(135, "")

End Sub

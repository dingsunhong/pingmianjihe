VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "认证对话框"
   ClientHeight    =   4380
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2587.849
   ScaleMode       =   0  'User
   ScaleWidth      =   5859.022
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   2040
      TabIndex        =   10
      Text            =   "000-000000"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确   定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4680
      TabIndex        =   2
      Top             =   3240
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取  消"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4680
      TabIndex        =   3
      Top             =   3720
      Width           =   1140
   End
   Begin VB.Label LabelUserName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "软件序列号:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   720
      TabIndex        =   6
      Top             =   2640
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "注意事项:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label lblLabels 
      Caption         =   "计算机身份证:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   3240
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      Caption         =   "软件认证码:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   3720
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Dim data_con2 As String
    '设置全局变量为 false
    '不提示失败的登录
 protect_data.serial_no = Text3.text
 If protect_data.install_statue = "S" Then
  If Abs(DateDiff("d", protect_data.install_date, Date)) > 15 Then '超过15天
    '未认证,超过认证期,成为使用版
        protect_data.install_statue = "F" '试用
         '设置使用版
        Call ExitProgram
'       Call set_enabled(False)
  Else
         Call set_enabled(True)
         LoginSucceeded = False
         MDIForm1.Timer2.Enabled = True
         Unload Me
         yes_Id = True
  End If
 ElseIf protect_data.install_statue = "F" Then
  '使用版,再认证
      Call ExitProgram
      'Call set_enabled(False)
End If
End Sub

Private Sub cmdOK_Click()
Dim data_con2 As String
Dim dir$
Dim pass_word As String
If event_statue = initial_condition Then
If Len(Text3.text) > 10 Then
 Text3.text = Mid$(Text3.text, 1, 10) ' 读出10个字符
ElseIf Len(Text3.text) < 10 Then '少于10个字符
       MsgBox LoadResString_(1845, ""), , LoadResString_(1850, "")
'         MsgBox "产品序列号少于10位,无效，请重新输入!", , LoadResString_(267)
Exit Sub
End If
protect_data.serial_no = Text3.text '设置产品序列号
pass_word = cal_password(protect_data.serial_no, protect_data.computer_id) '计算密码1
    '检查正确的密码
 If Trim(Text1.text) + Trim(Text2.text) + Trim(Text4.text) = _
                    pass_word Then '正确
            protect_data.pass_word = pass_word '设置密码
        If protect_data.input_pass_word_time < 900 Then '输入密码次树<900
          LoginSucceeded = True
           '认证成功
            protect_data.install_statue = "T"
        End If
      MDIForm1.Timer2.Enabled = True
     Unload Me
  Else
      pass_word = cal_password_1(protect_data.serial_no, protect_data.computer_id) '计算密码2
        If Trim(Text1.text) + Trim(Text2.text) + Trim(Text4.text) = _
                       pass_word Then '正确校园版
          protect_data.pass_word = pass_word
           If protect_data.input_pass_word_time < 900 Then
            LoginSucceeded = True
             '认证成功
            protect_data.install_statue = "T"
         End If
         MDIForm1.Timer2.Enabled = True
       Unload Me
      Else
       If protect_data.id = "31304619" Then
        protect_data.input_pass_word_time = protect_data.input_pass_word_time + 1
             MsgBox LoadResString_(1855, ""), , LoadResString_(1850, "")
       Else
         protect_data.id = "31304619"
         protect_data.input_pass_word_time = 1
       End If
         Text1.SetFocus
          SendKeys "{Home}+{End}"
      End If
    End If
    yes_Id = True
  ElseIf event_statue = exit_program Then
         event_statue = initial_condition
    Label4.visible = True
    lblLabels(0).visible = True
    lblLabels(1).visible = True
    LabelUserName.visible = True
    Text1.visible = True
    Text2.visible = True
    Text3.visible = True
    Text4.visible = True
    frmLogin.Label1 = LoadResString_(1835, "")
    cmdOK.Caption = LoadResString_(3940, "")
    cmdCancel.Caption = LoadResString_(3945, "")
  End If
End Sub

Private Sub Form_Load()
MDIForm1.WindowState = 2
Wenti_form.width = MDIForm1.width / 2
Draw_form.width = MDIForm1.width / 2
frmLogin.Enabled = True
Me.Caption = LoadResString_(500, "")
Label1.Caption = ""
Label2.Caption = LoadResString_(495, "")
Label4.Caption = LoadResString_(2265, "")
lblLabels(0).Caption = LoadResString_(2270, "")
lblLabels(1).Caption = LoadResString_(4055, "")
cmdOK.Caption = LoadResString_(3940, "")
cmdCancel.Caption = LoadResString_(3945, "")
If Len(Trim(protect_data.serial_no)) < 10 Then
   protect_data.serial_no = "0000000000"
End If
End Sub


Private Sub Text1_Change()
If Len(Text1.text) >= 4 Then
 Text1.text = Mid$(Text1.text, 1, 4)
 Text1.SetFocus
 Text2.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii < 0 Or KeyAscii > 127 Then
   KeyAscii = 0
End If
End Sub

Private Sub Text2_Change()
If Len(Text2.text) >= 4 Then
 Text2.text = Mid$(Text2.text, 1, 4)
   
  Text4.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
   KeyAscii = 0
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii < 0 Or KeyAscii > 127 Then
   KeyAscii = 0
End If
End Sub

Private Sub Text4_Change()
If Len(Text4.text) >= 4 Then
 Text4.text = Mid$(Text4.text, 1, 4)
  Me.cmdOK.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii < 0 Or KeyAscii > 127 Then
   KeyAscii = 0
End If
End Sub

Public Sub ExitProgram()
 If event_statue = initial_condition Then
    event_statue = exit_program
    Label4.visible = False
    lblLabels(0).visible = False
    lblLabels(1).visible = False
    LabelUserName.visible = False
    Text1.visible = False
    Text2.visible = False
    Text3.visible = False
    Text4.visible = False
    cmdOK.Caption = LoadResString_(560, "")
    cmdCancel.Caption = LoadResString_(135, "")
    Label1.Caption = LoadResString_(535, "")
 ElseIf event_statue = exit_program Then
    'End
    MDIForm1.Timer2.Enabled = True
    event_statue = exit_program
    Unload Me
    frmAbout.Show
 End If
End Sub

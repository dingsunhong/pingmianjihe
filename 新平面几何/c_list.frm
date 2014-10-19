VERSION 5.00
Begin VB.Form clist_form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form6"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   240
      Width           =   6135
   End
   Begin VB.CommandButton Confirm 
      Caption         =   "确定"
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Exit 
      Caption         =   "退出"
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Three3 
      Caption         =   "第三册"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Two2 
      Caption         =   "第二册"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton One1 
      Caption         =   "第一册"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.ListBox List6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   3840
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   6135
   End
End
Attribute VB_Name = "clist_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim volume As Variant


'Private Sub cancel_Click()    ' “取消”命令按钮的功能代码
'  clist_form.List6.Clear           ' 对List6列表框全部清屏
'  clist_form.One1.Enabled = True
'  clist_form.Two2.Enabled = True
'  clist_form.Three3.Enabled = True
'End Sub

Private Sub Confirm_Click()    '“确定”命令按钮的功能代码
 clist_form.List6.Refresh           '   刷新List6 列表框控件
 clist_form.One1.Enabled = True     '  “第一册”命令按钮功能有效
 clist_form.Two2.Enabled = True     '  “第二册”命令按钮功能有效
 clist_form.Three3.Enabled = True   '  “第三册”命令按钮功能有效
'clist_form.cancel.Enabled = True   '   “取消”命令按钮功能有效
 clist_form.Exit.Enabled = True     '   “退出”命令按钮功能有效
End Sub

Private Sub exit_Click()         '“退出”命令按钮功能代码
  End
End Sub

Private Sub One1_Click()       '“第一册”命令按钮的功能代码
clist_form.List6.Clear              ' 对List6列表框全部清屏
clist_form.One1.Enabled = True      '“第一册”命令按钮功能有效
clist_form.Two2.Enabled = True      ' “第二册”命令按钮功能有效
clist_form.Three3.Enabled = True    '  “第三册”命令按钮功能有效
'clist_form.cancel.Enabled = True   '   “取消”命令按钮功能有效
clist_form.Exit.Enabled = True      '   “退出”命令按钮功能有效
clist_form.Confirm.Enabled = True   '  “确定 ”命令按钮功能有效
 Call set_gate1(1)             '调用显示第一册目录的过程函数“set_gate1(1)”
End Sub


Private Sub Text6_Click()      '“List6”文本框的功能代码
clist_form.Text6.Refresh            '对“List6”文本框控件刷新
clist_form.Text6.SetFocus
'clist_form.Text6.Visible
End Sub                        '活动文本框的具体功能需要添加语句

Private Sub Three3_Click()     '“第三册”命令按钮的功能代码
clist_form.List6.Clear              ' 对List6列表框全部清屏
clist_form.One1.Enabled = True      '  “第一册”命令按钮功能有效
clist_form.Two2.Enabled = True      ' “第二册”命令按钮功能有效
clist_form.Three3.Enabled = True    '  “第三册”命令按钮功能有效
'clist_form.cancel.Enabled = True   '   “取消”命令按钮功能有效
clist_form.Exit.Enabled = True      '   “退出”命令按钮功能有效
clist_form.Confirm.Enabled = True   '  “确定 ”命令按钮功能有效
 Call set_gate3(3)             '调用显示第三册目录的过程函数“set_gate3(3)”
End Sub

Private Sub Two2_Click()       '“第二册”命令按钮的功能代码
clist_form.List6.Clear              ' 对List6列表框全部清屏
clist_form.Two2.Enabled = True      ' “第二册”命令按钮功能有效
clist_form.One1.Enabled = True      '  “第一册”命令按钮功能有效
clist_form.Three3.Enabled = True    '  “第一册”命令按钮功能有效
'clist_form.cancel.Enabled = True   '“取消”命令按钮功能有效
clist_form.Exit.Enabled = True           ' “退出”命令按钮功能有效
clist_form.Confirm.Enabled = True   '   “确定”命令按钮功能有效
 Call set_gate2(2)             '调用显示第二册目录的过程函数“set_gate2(2)”
 
End Sub

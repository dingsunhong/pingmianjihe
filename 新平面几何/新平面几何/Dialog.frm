VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "软件认证对话框 "
   ClientHeight    =   5175
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6240
   ControlBox      =   0   'False
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取  消"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确  定"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "请输入认证密码:"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2535
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "您的计算机的身份证号: "
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Me.visible = False
End Sub

Private Sub OKButton_Click()
Dim data_con As String
data_con = Trim(Me.Text1.text)
If protect_data.pass_word = Trim(Me.Text1.text) Then
    protect_data.pass_word1 = data_con
Else
 Me.Text1.text = "认证密码错,请重输!"
End If
End Sub

Private Sub OKButton_KeyPress(KeyAscii As Integer)
If Me.Text1.text = "认证密码错,请重输!" Then
   Me.Text1.text = ""
End If
If Len(Text1.text) < 6 Then
If KeyAscii >= 48 And KeyAscii < 58 Then
 Me.Text1.text = Me.Text1.text + Chr(KeyAscii)
End If
End If
End Sub

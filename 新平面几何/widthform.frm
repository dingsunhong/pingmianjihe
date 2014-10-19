VERSION 5.00
Begin VB.Form widthform 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "设置画图板属性"
   ClientHeight    =   3036
   ClientLeft      =   2868
   ClientTop       =   1716
   ClientWidth     =   5952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3036
   ScaleWidth      =   5952
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   8
      Top             =   1440
      Width           =   1932
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "辅助线色："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1692
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   7
      Top             =   960
      Width           =   1932
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "结论颜色："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1932
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   6
      Top             =   480
      Width           =   1932
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "条件颜色："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1692
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   5
      Top             =   0
      Width           =   1932
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "线宽："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1692
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取 消"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确  定"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   1920
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   2
      Top             =   0
      Width           =   4092
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "例 样"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   372
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1092
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   1
      Top             =   2400
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   360
      Picture         =   "widthform.frx":0000
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   251
      TabIndex        =   0
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   192
      Left            =   960
      Picture         =   "widthform.frx":112A
      Top             =   2160
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "widthform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Call set_draw_width
'Call draw_again1(Draw_form)
'line_width = temp_line_width
'condition_color = temp_condition_color
'conclusion_color = temp_conclusion_color
'fill_color = temp_fill_color
'Draw_form.DrawWidth = line_width
'Call init_set
'Call init_color
widthform.Hide
'Call draw_again1(Draw_form)
End Sub


Private Sub Command2_Click()
'temp_line_color = line_color
temp_condition_color = condition_color
temp_conclusion_color = conclusion_color
temp_line_width = line_width
Call init_set
widthform.Hide
End Sub

Private Sub Form_Load()
widthform.Caption = LoadResString_(4200, "")
Label1.Caption = LoadResString_(4205, "")
Label2.Caption = LoadResString_(4210, "")
Label3.Caption = LoadResString_(4215, "")
Label4.Caption = LoadResString_(4220, "")
Label5.Caption = LoadResString_(4225, "")
Command1.Caption = LoadResString_(3940, "")
Command2.Caption = LoadResString_(135, "")
End Sub

Private Sub Label1_Click()
Call init_color
Label1.BackColor = QBColor(14)
width_set_statue = 1
End Sub

Private Sub Label2_Click()
Call init_color
Label2.BackColor = QBColor(14)
width_set_statue = 2

End Sub

Private Sub Label3_Click()
Call init_color
Label3.BackColor = QBColor(14)
width_set_statue = 3
End Sub
Private Sub Label4_Click()
Call init_color
Label4.BackColor = QBColor(14)
width_set_statue = 4
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim t%, i%
t% = Int(Y)
i% = Int(X)
Dim c As Byte
If Int(X) < 30 And width_set_statue = 1 Then
If t% > 6 And t% < 10 Then
temp_line_width = 1
ElseIf t% > 10 And t% < 15 Then
temp_line_width = 2
ElseIf t% > 15 And t% < 21 Then
temp_line_width = 3
ElseIf t% > 21 And t% < 28 Then
temp_line_width = 4
ElseIf t% > 28 And t% < 37 Then
temp_line_width = 5
End If
ElseIf width_set_statue > 1 Then
i% = Int(X)
 If t% > 7 And t% < 20 Then
   If i% > 27 And i% < 49 Then
    c = 0
   ElseIf i% > 48 And i% < 70 Then
    c = 1
   ElseIf i% > 69 And i% < 91 Then
    c = 2
   ElseIf i% > 90 And i% < 112 Then
    c = 3
   ElseIf i% > 111 And i% < 133 Then
    c = 4
   ElseIf i% > 132 And i% < 154 Then
    c = 5
   ElseIf i% > 153 And i% < 175 Then
    c = 6
   ElseIf i% > 174 And i% < 196 Then
    c = 7
   End If
 ElseIf t% > 20 And t% < 33 Then
   If i% > 27 And i% < 49 Then
    c = 8
   ElseIf i% > 48 And i% < 70 Then
    c = 9
   ElseIf i% > 69 And i% < 91 Then
    c = 10
   ElseIf i% > 90 And i% < 112 Then
    c = 11
   ElseIf i% > 111 And i% < 133 Then
    c = 12
   ElseIf i% > 132 And i% < 154 Then
    c = 13
   ElseIf i% > 153 And i% < 175 Then
    c = 14
   ElseIf i% > 174 And i% < 196 Then
    c = 15
   End If
End If
End If
If width_set_statue = 1 Then
Call draw_arow
Call draw_condition_color
Call draw_conclusion
Call draw_fill_color
Call linewidth
Call draw_sample
ElseIf width_set_statue = 2 And c > 0 Then
temp_condition_color = c
Call draw_condition_color
Call draw_sample
ElseIf width_set_statue = 3 And c > 0 Then
temp_conclusion_color = c
Call draw_conclusion
Call draw_sample
ElseIf width_set_statue = 4 And c > 0 Then
temp_fill_color = c
Call draw_fill_color
End If
End Sub


Private Sub picture2_Click()
Call init_color
Label2.BackColor = QBColor(14)
width_set_statue = 2

End Sub

Private Sub Picture3_Click()
Call init_set
End Sub


Private Sub Picture4_Click()
Call init_color
Label1.BackColor = QBColor(14)
width_set_statue = 1
End Sub


Private Sub Picture5_Click()
Call init_color
Label2.BackColor = QBColor(14)
width_set_statue = 2
End Sub


Private Sub Picture6_Click()
Call init_color
Label3.BackColor = QBColor(14)
width_set_statue = 3
End Sub


Private Sub Picture7_Click()
Call init_color
Label4.BackColor = QBColor(14)
width_set_statue = 4
End Sub



'

Private Sub Picture8_Click()

End Sub



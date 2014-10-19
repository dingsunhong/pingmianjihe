VERSION 5.00
Begin VB.Form Print_Form 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "打印编辑"
   ClientHeight    =   10332
   ClientLeft      =   216
   ClientTop       =   660
   ClientWidth     =   8520
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.4
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10332
   ScaleWidth      =   8520
   StartUpPosition =   2  '屏幕中心
   Begin VB.VScrollBar VScroll1 
      Height          =   10100
      Left            =   8280
      Max             =   100
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      Left            =   0
      Max             =   100
      TabIndex        =   2
      Top             =   10100
      Width           =   8295
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   50000
      Left            =   120
      MousePointer    =   3  'I-Beam
      ScaleHeight     =   50004
      ScaleWidth      =   12000
      TabIndex        =   1
      Top             =   0
      Width           =   12000
      Begin VB.TextBox Text1 
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   360
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   480
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "页眉:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30000
      Left            =   0
      ScaleHeight     =   30000
      ScaleWidth      =   12000
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "后一页"
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
         Left            =   7000
         TabIndex        =   5
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "前一页"
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
         Left            =   5640
         TabIndex        =   4
         Top             =   5760
         Width           =   1095
      End
   End
   Begin VB.Menu page_and_note 
      Caption         =   "页眉批注"
      Begin VB.Menu page_note 
         Caption         =   "更改页眉"
      End
      Begin VB.Menu st_page 
         Caption         =   "起始页码"
      End
      Begin VB.Menu note 
         Caption         =   "批注"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu preview 
      Caption         =   "打印预览"
   End
   Begin VB.Menu print1 
      Caption         =   "打　印"
   End
   Begin VB.Menu exit 
      Caption         =   "退  出"
   End
End
Attribute VB_Name = "Print_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim p_left(1) As Integer
Dim p_top(1) As Integer
Dim t_p_no%
Dim caret_display_or_no As Boolean
Dim s_x&, s_y&, s_x1&, s_y1&
Dim mouse_action As Boolean
Dim text_start_x%, text_start_y%
Dim caret_x%, caret_y%

Private Sub Command1_Click()
If t_p_no% > 1 Then
Print_Form.Picture2.Cls
t_p_no% = t_p_no% - 1
Call preview_(t_p_no%)
   Call set_command_box(t_p_no%)
End If
End Sub

Private Sub Command2_Click()
If t_p_no% < page_no% Then
Print_Form.Picture2.Cls
t_p_no% = t_p_no% + 1
Call preview_(t_p_no%)
   Call set_command_box(t_p_no%)
End If
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mouse_action = False
End Sub

Private Sub HScroll1_Change()
left0% = HScroll1.value * (Print_Form.width - 14400) / 100
Picture1.left = p_left(0) + left0%
'Picture2.Left = p_left(1) + left0%
End Sub



Private Sub page_note_Click()
Me.Label1.Caption = LoadResString_(2040, "")
Me.Label1.visible = True
Me.Text1.visible = True
Me.Picture1.Cls
Me.Text1.text = page_note_string
End Sub

Private Sub preview_Click()
If Me.preview.Caption = LoadResString_(130, "") And C_display_wenti.m_page_n > 0 Then
Print_Form.Picture1.visible = False
Print_Form.Picture2.visible = True
t_p_no% = 1
 Call preview_(1)
  Me.preview.Caption = LoadResString_(2045, "")
   Call set_command_box(1)
Else
 Print_Form.Picture1.visible = True
  Print_Form.Picture2.visible = False
  Me.preview.Caption = LoadResString_(130, "")
End If
End Sub

Private Sub print1_Click()
'On Error GoTo print1_error
Printer.font = LoadResString_(2050, "")
Call C_display_wenti.m_Print_wenti(Printer, 1, 2, 0, True)
Exit Sub
print1_error:
MsgBox LoadResString_(2255, ""), vbOKOnly
End Sub

Private Sub st_page_Click()
Me.Label1.Caption = LoadResString_(2055, "")
Me.Label1.visible = True
Me.Text1.visible = True
Me.Picture1.Cls
Me.Text1.text = str(start_page_no% + 1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Me.Label1.visible = False
 Me.Text1.visible = False
 If Me.Label1.Caption = LoadResString_(2040, "") Then
 page_note_string = Me.Text1.text
 Else
 start_page_no% = val(Me.Text1.text) - 1
 End If
 Call Me.Picture1.Cls
 Call C_display_wenti.m_Print_wenti(Print_Form.Picture1, 0, 2, 0, True)
End If
End Sub

Private Sub VScroll1_Change()
top0% = VScroll1.value * Print_Form.Picture1.Height / 100
Picture1.top = p_top(0) - top0%
'Picture2.Top = p_top(1) + top0%

End Sub



Private Sub Picture1_KeyPress(KeyAscii As Integer)
Dim X%, Y%
If Print_Form.CurrentX < text_start_x% Then
 Print_Form.CurrentX = text_start_x%
End If

last_edit_char = last_edit_char + 1
ReDim Preserve E_char(last_edit_char) As edit_char_type
E_char(last_edit_char).ch = Chr(KeyAscii)
E_char(last_edit_char).pos.X = Print_Form.CurrentX
E_char(last_edit_char).pos.Y = Print_Form.CurrentY
If Picture1.top - Print_Form.CurrentY > 0 And _
       Picture2.top - Print_Form.CurrentY > 0 And _
        (Picture1.top - Print_Form.CurrentY < 500 Or _
          Picture1.top - Print_Form.CurrentY < 500) Then
Picture1.top = Picture1.top + 500
Picture2.top = Picture2.top + 500
p_top(0) = p_top(0) + 500
p_top(1) = p_top(1) + 500
ElseIf Picture1.top - Print_Form.CurrentY > 0 And _
 Picture1.top - Print_Form.CurrentY < 500 Then
Picture1.top = Picture1.top + 500
p_top(0) = p_top(0) + 500
ElseIf Picture2.top - Print_Form.CurrentY > 0 And _
   Picture2.top - Print_Form.CurrentY < 500 Then
Picture2.top = Picture2.top + 500
p_top(1) = p_top(1) + 500
End If
   
If Print_Form.CurrentX > 7000 Then
 X% = Print_Form.CurrentX
 Y% = Print_Form.CurrentY
 Print_Form.Line (caret_x, caret_y + 300)- _
 (caret_x + 200, caret_y + 300), QBColor(15)
 Print_Form.CurrentX = X%
 Print_Form.CurrentY = Y%
 Print_Form.Print Chr(KeyAscii)
 caret_x% = text_start_x%
 caret_y% = Print_Form.CurrentY
 'Call display_caret
 Print_Form.Line (caret_x, caret_y + 300)- _
   (caret_x + 200, caret_y + 300), QBColor(12)
 Print_Form.CurrentX = caret_x%
 Print_Form.CurrentY = caret_y%
 
Else
'MDIForm1.Timer1.Enabled = False
 X% = Print_Form.CurrentX
 Y% = Print_Form.CurrentY

 Print_Form.Line (caret_x, caret_y + 300)- _
 (caret_x + 200, caret_y + 300), QBColor(15)
 Print_Form.CurrentX = X%
 Print_Form.CurrentY = Y%

 Print_Form.Print Chr(KeyAscii);
 caret_x% = Print_Form.CurrentX
 caret_y% = Print_Form.CurrentY
 'Call display_caret
 Print_Form.Line (caret_x, caret_y + 300)- _
   (caret_x + 200, caret_y + 300), QBColor(12)
 Print_Form.CurrentX = caret_x%
 Print_Form.CurrentY = caret_y%

End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
s_x& = Int(X)
s_y& = Int(Y)
s_x1 = Picture1.left
s_y1& = Picture1.top
End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Picture1.top = Picture1.top + Int(Y) - s_y&
Picture1.left = Picture1.left + Int(X) - s_x&

'Call change_edit
End If

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
p_top(0) = p_top(0) + Picture1.top - s_y1&
p_left(0) = p_left(0) + Picture1.left - s_x&
End If
End Sub


Private Sub picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
s_x& = Int(X)
s_y& = Int(Y)
s_x1 = Picture2.left
s_y1& = Picture2.top

End If
End Sub

Private Sub picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Picture2.top = Picture2.top + Int(Y) - s_y&
Picture2.left = Picture2.left + Int(X) - s_x&
'Call change_edit
End If
End Sub

Private Sub picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
p_top(1) = p_top(1) + Picture2.top - s_y1&
p_left(1) = p_left(1) + Picture2.left - s_x1&
End If
End Sub


Public Sub display_caret()
Dim X%, Y%
Print_Form.CurrentX = X%
Print_Form.CurrentY = Y%
MDIForm1.Timer1.Enabled = True
Do
Do
DoEvents
Loop Until time_act = True
If caret_display_or_no Then
 caret_display_or_no = False
 Print_Form.Line (caret_x, caret_y + 300)- _
 (caret_x + 200, caret_y + 300), QBColor(12)
Else
 caret_display_or_no = True
  Print_Form.Line (caret_x, caret_y + 300)- _
 (caret_x + 200, caret_y + 300), QBColor(15)
End If
time_act = False
Loop Until MDIForm1.Timer1.Enabled = False
End Sub

Public Sub preview_(p_no%)
  Call C_display_wenti.m_Print_wenti(Print_Form.Picture1, 0, 1, p_no% - 1, False)
   Print_Form.Picture2.Line (2000, 500)-(5200, 5300), , B
   Print_Form.Picture2.Line (5216, 520)-(5216, 5340), QBColor(7)
   Print_Form.Picture2.Line (5232, 520)-(5232, 5340), QBColor(7)
   Print_Form.Picture2.Line (5248, 520)-(5248, 5340), QBColor(7)
   Print_Form.Picture2.Line (2020, 5310)-(5248, 5310), QBColor(7)
   Print_Form.Picture2.Line (2020, 5320)-(5248, 5320), QBColor(7)
   Print_Form.Picture2.Line (2020, 5330)-(5248, 5330), QBColor(7)
   Print_Form.Picture2.Line (2020, 5340)-(5248, 5340), QBColor(7)
   Print_Form.Picture2.CurrentX = 3200 ' j_x% + 1200
   Print_Form.Picture2.CurrentY = 5500 'j_y% + 5000
   Print_Form.Picture2.Print "(" + str(start_page_no% + p_no%) + ")"
 Call StretchBlt(Print_Form.Picture2.hdc, 151, 58, 230, 400, _
       Print_Form.Picture1.hdc, 0, 0, 700, 1200, &H8800C6) '将打印结果映射到打印预览窗口
End Sub

Public Sub set_command_box(t_p_no%)
If page_no% < 2 Then
Me.Command1.Enabled = False
Me.Command2.Enabled = False
Else
If t_p_no% = 1 Then
Me.Command1.Enabled = False
Me.Command2.Enabled = True
ElseIf t_p_no% = page_no% Then
Me.Command1.Enabled = True
Me.Command2.Enabled = False
Else
Me.Command1.Enabled = True
Me.Command2.Enabled = True
End If
End If
End Sub

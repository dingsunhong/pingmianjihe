VERSION 5.00
Begin VB.Form P_Form 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "´òÓ¡±à¼­"
   ClientHeight    =   6480
   ClientLeft      =   210
   ClientTop       =   660
   ClientWidth     =   8520
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ËÎÌå"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6480
   ScaleWidth      =   8520
   Begin VB.VScrollBar VScroll1 
      Height          =   6255
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
      Top             =   6240
      Width           =   8295
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   90000
      Left            =   120
      MousePointer    =   3  'I-Beam
      ScaleHeight     =   90000
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
         Caption         =   "Ò³Ã¼:"
         BeginProperty Font 
            Name            =   "ËÎÌå"
            Size            =   10.5
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   90000
      Left            =   0
      ScaleHeight     =   90000
      ScaleWidth      =   12000
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12000
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ºóÒ»Ò³"
         BeginProperty Font 
            Name            =   "ËÎÌå"
            Size            =   10.5
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
         Caption         =   "Ç°Ò»Ò³"
         BeginProperty Font 
            Name            =   "ËÎÌå"
            Size            =   10.5
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
      Caption         =   "Ò³Ã¼Åú×¢"
      Begin VB.Menu page_note 
         Caption         =   "¸ü¸ÄÒ³Ã¼"
      End
      Begin VB.Menu st_page 
         Caption         =   "ÆðÊ¼Ò³Âë"
      End
      Begin VB.Menu note 
         Caption         =   "Åú×¢"
      End
   End
   Begin VB.Menu preview 
      Caption         =   "´òÓ¡Ô¤ÀÀ"
   End
   Begin VB.Menu print1 
      Caption         =   "´ò¡¡Ó¡"
   End
   Begin VB.Menu exit 
      Caption         =   "ÍË  ³ö"
   End
End
Attribute VB_Name = "P_Form"
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
P_Form.Picture2.Cls
t_p_no% = t_p_no% - 1
Call preview_(2000, 500, 0, 1065 * (t_p_no% - 1), t_p_no%)
   Call set_command_box(t_p_no%)
End If
End Sub

Private Sub Command2_Click()
If t_p_no% < page_no% Then
P_Form.Picture2.Cls
t_p_no% = t_p_no% + 1
Call preview_(2000, 500, 0, 1065 * (t_p_no% - 1), t_p_no%)
   Call set_command_box(t_p_no%)
End If
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
mouse_action = False
End Sub

Private Sub HScroll1_Change()
left0% = HScroll1.value * (P_Form.Width - 14400) / 100
Picture1.Left = p_left(0) + left0%
'Picture2.Left = p_left(1) + left0%
End Sub



Private Sub page_note_Click()
Me.Label1.Caption = "Ò³Ã¼£º"
Me.Label1.visible = True
Me.Text1.visible = True
Me.Picture1.Cls
Me.Text1.text = page_note_string
End Sub

Private Sub preview_Click()
If Me.preview.Caption = "´òÓ¡Ô¤ÀÀ" And page_no% > 0 Then
P_Form.Picture1.visible = False
P_Form.Picture2.visible = True
t_p_no% = 1
 Call preview_(2000, 500, 0, 0, 1)
  Me.preview.Caption = "¹Ø±ÕÔ¤ÀÀ"
   Call set_command_box(1)
Else
 P_Form.Picture1.visible = True
  P_Form.Picture2.visible = False
  Me.preview.Caption = "´òÓ¡Ô¤ÀÀ"
End If
End Sub

Private Sub print1_Click()
Call P_wenti(Printer, 1, 2)
End Sub

Private Sub st_page_Click()
Me.Label1.Caption = "ÆðÊ¼Ò³Âë£º"
Me.Label1.visible = True
Me.Text1.visible = True
Me.Picture1.Cls
Me.Text1.text = str(start_page_no% + 1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Me.Label1.visible = False
 Me.Text1.visible = False
 If Me.Label1.Caption = "Ò³Ã¼£º" Then
 page_note_string = Me.Text1.text
 Else
 start_page_no% = Val(Me.Text1.text) - 1
 End If
 Call Me.Picture1.Cls
 Call P_wenti(P_Form.Picture1, 0, 2)
End If
End Sub

Private Sub VScroll1_Change()
top0% = VScroll1.value * P_Form.Picture1.Height / 100
Picture1.Top = p_top(0) - top0%
'Picture2.Top = p_top(1) + top0%

End Sub



Private Sub Picture1_KeyPress(KeyAscii As Integer)
Dim x%, y%
If P_Form.CurrentX < text_start_x% Then
 P_Form.CurrentX = text_start_x%
End If

last_edit_char = last_edit_char + 1
ReDim Preserve E_char(last_edit_char) As edit_char_type
E_char(last_edit_char).ch = Chr(KeyAscii)
E_char(last_edit_char).pos.x = P_Form.CurrentX
E_char(last_edit_char).pos.y = P_Form.CurrentY
If Picture1.Top - P_Form.CurrentY > 0 And _
       Picture2.Top - P_Form.CurrentY > 0 And _
        (Picture1.Top - P_Form.CurrentY < 500 Or _
          Picture1.Top - P_Form.CurrentY < 500) Then
Picture1.Top = Picture1.Top + 500
Picture2.Top = Picture2.Top + 500
p_top(0) = p_top(0) + 500
p_top(1) = p_top(1) + 500
ElseIf Picture1.Top - P_Form.CurrentY > 0 And _
 Picture1.Top - P_Form.CurrentY < 500 Then
Picture1.Top = Picture1.Top + 500
p_top(0) = p_top(0) + 500
ElseIf Picture2.Top - P_Form.CurrentY > 0 And _
   Picture2.Top - P_Form.CurrentY < 500 Then
Picture2.Top = Picture2.Top + 500
p_top(1) = p_top(1) + 500
End If
   
If P_Form.CurrentX > 7000 Then
 x% = P_Form.CurrentX
 y% = P_Form.CurrentY
 P_Form.Line (caret_x, caret_y + 300)- _
 (caret_x + 200, caret_y + 300), QBColor(15)
 P_Form.CurrentX = x%
 P_Form.CurrentY = y%
 P_Form.Print Chr(KeyAscii)
 caret_x% = text_start_x%
 caret_y% = P_Form.CurrentY
 'Call display_caret
 P_Form.Line (caret_x, caret_y + 300)- _
   (caret_x + 200, caret_y + 300), QBColor(12)
 P_Form.CurrentX = caret_x%
 P_Form.CurrentY = caret_y%
 
Else
'MDIForm1.Timer1.Enabled = False
 x% = P_Form.CurrentX
 y% = P_Form.CurrentY

 P_Form.Line (caret_x, caret_y + 300)- _
 (caret_x + 200, caret_y + 300), QBColor(15)
 P_Form.CurrentX = x%
 P_Form.CurrentY = y%

 P_Form.Print Chr(KeyAscii);
 caret_x% = P_Form.CurrentX
 caret_y% = P_Form.CurrentY
 'Call display_caret
 P_Form.Line (caret_x, caret_y + 300)- _
   (caret_x + 200, caret_y + 300), QBColor(12)
 P_Form.CurrentX = caret_x%
 P_Form.CurrentY = caret_y%

End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
s_x& = Int(x)
s_y& = Int(y)
s_x1 = Picture1.Left
s_y1& = Picture1.Top
End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
Picture1.Top = Picture1.Top + Int(y) - s_y&
Picture1.Left = Picture1.Left + Int(x) - s_x&

'Call change_edit
End If

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
p_top(0) = p_top(0) + Picture1.Top - s_y1&
p_left(0) = p_left(0) + Picture1.Left - s_x&
End If
End Sub


Private Sub picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
s_x& = Int(x)
s_y& = Int(y)
s_x1 = Picture2.Left
s_y1& = Picture2.Top

End If
End Sub

Private Sub picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
Picture2.Top = Picture2.Top + Int(y) - s_y&
Picture2.Left = Picture2.Left + Int(x) - s_x&
'Call change_edit
End If
End Sub

Private Sub picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
p_top(1) = p_top(1) + Picture2.Top - s_y1&
p_left(1) = p_left(1) + Picture2.Left - s_x1&
End If
End Sub


Public Sub display_caret()
Dim x%, y%
P_Form.CurrentX = x%
P_Form.CurrentY = y%
MDIForm1.Timer1.Enabled = True
Do
Do
DoEvents
Loop Until time_act = True
If caret_display_or_no Then
 caret_display_or_no = False
 P_Form.Line (caret_x, caret_y + 300)- _
 (caret_x + 200, caret_y + 300), QBColor(12)
Else
 caret_display_or_no = True
  P_Form.Line (caret_x, caret_y + 300)- _
 (caret_x + 200, caret_y + 300), QBColor(15)
End If
time_act = False
Loop Until MDIForm1.Timer1.Enabled = False
End Sub

Public Sub preview_(j_x%, j_y%, s_x%, s_y%, p_no%)
   P_Form.Picture2.Line (j_x%, j_y%)-(j_x% + 3200, j_y% + 4800), , B
   P_Form.Picture2.Line (j_x% + 3216, j_y% + 20)-(j_x% + 3216, j_y% + 4840), QBColor(7)
   P_Form.Picture2.Line (j_x% + 3232, j_y% + 20)-(j_x% + 3232, j_y% + 4840), QBColor(7)
   P_Form.Picture2.Line (j_x% + 3248, j_y% + 20)-(j_x% + 3248, j_y% + 4840), QBColor(7)
   P_Form.Picture2.Line (j_x% + 20, j_y% + 4810)-(j_x% + 3248, j_y% + 4810), QBColor(7)
   P_Form.Picture2.Line (j_x% + 20, j_y% + 4820)-(j_x% + 3248, j_y% + 4820), QBColor(7)
   P_Form.Picture2.Line (j_x% + 20, j_y% + 4830)-(j_x% + 3248, j_y% + 4830), QBColor(7)
   P_Form.Picture2.Line (j_x% + 20, j_y% + 4840)-(j_x% + 3248, j_y% + 4840), QBColor(7)
   P_Form.Picture2.CurrentX = j_x% + 1200
   P_Form.Picture2.CurrentY = j_y% + 5000
   P_Form.Picture2.Print "(µÚ" + str(start_page_no% + p_no%) + "Ò³)"
 Call StretchBlt(P_Form.Picture2.hdc, j_x% / 18, j_y% / 13, 210, 315, _
       P_Form.Picture1.hdc, s_x%, s_y%, 600, 1000, &H8800C6)

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

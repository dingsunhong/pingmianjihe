VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form inform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "　信息库"
   ClientHeight    =   4215
   ClientLeft      =   1095
   ClientTop       =   1365
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   442
   Begin VB.VScrollBar VScroll1 
      Height          =   4215
      Left            =   3960
      Max             =   100
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1.12050e5
      Left            =   0
      ScaleHeight     =   7466
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      Begin ComctlLib.TreeView TreeView1 
         Height          =   4215
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7435
         _Version        =   327682
         LineStyle       =   1
         Style           =   6
         Appearance      =   1
         MousePointer    =   1
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
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   4200
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "inform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim old_wenti_no As Integer
Dim old_fontname As String
Dim old_fontsize As Integer



Private Sub Form_Load()
TreeView1.visible = inform_treeview_visible
VScroll1.visible = inform_picture_visible
old_fontname = inform.Picture1.FontName
old_fontsize = inform.Picture1.FontSize
Me.Caption = LoadResString_(955, "")
End Sub

Private Sub list1_Click()
Dim i%, n%
Dim date_type As Byte
Dim date_no%
Dim nod As Node
Dim ind As String
Dim temp_record As total_record_type
Last_hotpoint_of_theorem1 = 0
Erase Hotpoint_of_theorem1
inform.Picture1.Cls
inform.Picture1.top = 0
inform.VScroll1.value = 0
display_no = 0
'Call C_display_wenti1.init(inform.Picture1, wenti_condition_no)
old_wenti_no = C_display_wenti.m_last_input_wenti_no
If inform_type = eline_ And List1.ListIndex + 1 > last_conditions.last_cond(1).eline_no Then
  date_type = midpoint_
   date_no% = List1.ListIndex + 1 - last_conditions.last_cond(1).eline_no
ElseIf inform_type = polygon_ Then
 If 0 < List1.ListIndex + 1 And _
        List1.ListIndex + 1 <= last_conditions.last_cond(1).tixing_no Then
  date_type = tixing_
   date_no% = List1.ListIndex + 1
 ElseIf last_conditions.last_cond(1).tixing_no <= List1.ListIndex And _
       List1.ListIndex < last_conditions.last_cond(1).tixing_no Then
  date_type = equal_side_tixing_
   date_no% = List1.ListIndex + 1 - last_conditions.last_cond(1).tixing_no
 ElseIf last_conditions.last_cond(1).tixing_no <= List1.ListIndex And _
      List1.ListIndex < _
      last_conditions.last_cond(1).tixing_no + last_conditions.last_cond(1).parallelogram_no Then
  date_type = parallelogram_
   date_no% = List1.ListIndex + 1 - last_conditions.last_cond(1).tixing_no
 ElseIf last_conditions.last_cond(1).tixing_no + last_conditions.last_cond(1).parallelogram_no <= _
     List1.ListIndex And List1.ListIndex < _
      last_conditions.last_cond(1).tixing_no + last_conditions.last_cond(1).parallelogram_no + _
         last_conditions.last_cond(1).rhombus_no Then
  date_type = rhombus_
   date_no% = List1.ListIndex + 1 - _
     last_conditions.last_cond(1).tixing_no - last_conditions.last_cond(1).parallelogram_no
 ElseIf last_conditions.last_cond(1).tixing_no + last_conditions.last_cond(1).parallelogram_no + last_conditions.last_cond(1).rhombus_no _
    <= List1.ListIndex And List1.ListIndex < _
      last_conditions.last_cond(1).tixing_no + last_conditions.last_cond(1).parallelogram_no + _
         last_conditions.last_cond(1).rhombus_no + last_conditions.last_cond(1).long_squre_no Then
  date_type = long_squre_
   date_no% = List1.ListIndex + 1 - _
     last_conditions.last_cond(1).tixing_no - last_conditions.last_cond(1).parallelogram_no - _
       last_conditions.last_cond(1).rhombus_no
  ElseIf last_conditions.last_cond(1).tixing_no + last_conditions.last_cond(1).parallelogram_no + _
       last_conditions.last_cond(1).rhombus_no + last_conditions.last_cond(1).long_squre_no <= List1.ListIndex Then
   n% = 0
    For i% = 1 To last_conditions.last_cond(1).epolygon_no
     If epolygon(i%).data(0).p.total_v = 4 Then
       If n% = List1.ListIndex - last_conditions.last_cond(1).tixing_no - _
         last_conditions.last_cond(1).parallelogram_no - last_conditions.last_cond(1).rhombus_no - last_conditions.last_cond(1).long_squre_no Then
          date_type = epolygon_
            date_no% = i%
         ' Call set_display_string_no(epolygon_, i%)
       End If
      n% = n% + 1
      End If
     Next i%
   End If
ElseIf inform_type = angle_value_ Then
 date_type = angle3_value_
  date_no% = angle_value.av_no(List1.ListIndex + 1).no
ElseIf inform_type = Rangle_ Then
 date_type = angle3_value_
  date_no% = angle_value_90.av_no(List1.ListIndex + 1).no
ElseIf inform_type = eangle_ Then
 date_type = angle3_value_
  date_no% = Deangle.av_no(List1.ListIndex + 1).no
ElseIf inform_type = angle_relation_ Then
 date_type = angle3_value_
  date_no% = angle_relation.av_no(List1.ListIndex + 1).no
ElseIf inform_type = two_angle_value_sum_ Then
 date_type = angle3_value_
  date_no% = two_angle_value_sum.av_no(List1.ListIndex + 1).no
ElseIf inform_type = two_angle_180_ Then
 date_type = angle3_value_
  date_no% = Two_angle_value.av_no(List1.ListIndex + 1).no
ElseIf inform_type = angle2_right Then
 date_type = angle3_value_
  date_no% = two_angle_value_90.av_no(List1.ListIndex + 1).no
ElseIf inform_type = angle3_value_ Then
 date_type = angle3_value_
  date_no% = three_angle_value.av_no(List1.ListIndex + 1).no
Else
  date_type = inform_type
   date_no% = List1.ListIndex + 1
End If
Call draw_inform(0, date_type, date_no%)
If TreeView1.visible = False Then
Call C_display_wenti1.init(inform.Picture1, wenti_condition_no)
Call set_display_string_no(date_type, date_no%, 0, 0)
Call set_display_string(False, 0, 1, 0, True)
inform.Picture1.CurrentX = 0
inform.Picture1.CurrentY = 0
inform.Picture1.Print LoadResString_(1235, "")
 display_no = 0
Else
Call create_treeview(date_type, date_no%, TreeView1, TreeView1.top)
'TreeView1.Nodes.Clear
'last_node_index = 1
'ind = "node" & CStr(last_node_index)
'Set nod = TreeView1.Nodes.Add(, , ind, _
            set_display_string0(date_type, date_no%, 0, False, False, 0, 1, 0))
'ReDim Preserve cond_no(nod.index) As condition_no_type
'cond_no(nod.index).ty = date_type
'cond_no(nod.index).no = date_no%
'Call record_no(date_type, date_no%, temp_record, True, 0, 0)
'  If temp_record.record_data.data0.condition_data.condition_no > 0 And temp_record.record_data.data0.condition_data.condition_no < 9 Then
'  For i% = 1 To temp_record.record_data.data0.condition_data.condition_no
'  Call add_node(TreeView1, ind, _
'       temp_record.record_data.data0.condition_data.condition(i%).ty, temp_record.record_data.data0.condition_data.condition(i%).no)
'  Next i%
' End If
End If
End Sub

Private Sub List1_LostFocus()
'will disappear this form
'Me.Hide
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i%, n%
Call C_display_wenti1.get_MouseDown(Button, Int(X), Int(Y), 0)
 inform.SetFocus
End Sub
Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
 Me.Cls
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
Dim temp_record As total_record_type
Dim n%
If Node.index <= last_node_index Then
Call record_no(cond_no(Node.index).ty, cond_no(Node.index).no, temp_record, True, 0, 0)
n% = temp_record.record_data.data0.theorem_no
Call draw_inform(0, cond_no(Node.index).ty, cond_no(Node.index).no)
 If n% > 0 Then
 MDIForm1.StatusBar1.Panels(1).text = th_chose(n%).text
 Call UpdateWindow(MDIForm1.StatusBar1.hwnd)
End If
End If
End Sub

Private Sub VScroll1_Change()
inform.Picture1.top = (inform.Height - inform.Picture1.Height) * VScroll1.value / 100
End Sub



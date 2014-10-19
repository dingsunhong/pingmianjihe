VERSION 5.00
Begin VB.Form exam_form 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2424
   ClientLeft      =   2136
   ClientTop       =   1536
   ClientWidth     =   3372
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2424
   ScaleWidth      =   3372
   Visible         =   0   'False
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1128
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "exam_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim op_remove As Boolean
Private Sub Form_Load()
List1.Height = exam_form.Height - 2
List1.FontSize = 11.5
Call C_IO.add_exam_name_to_list(Me.List1)
End Sub
Private Sub List1_DblClick()
Dim i%
Dim t_name$
If path_and_file <> "" And save_statue = 1 Then
If MsgBox(LoadResString_(1285, "\\1\\" + path_and_file), 4, "", "", 0) = 6 Then
Call C_IO.save_prove_result(path_and_file)
End If
End If
Call C_IO.input_wenti_from_exam(List1.ListIndex + 1)  '
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ty%, i%, j%
Dim n%
If Button = 2 Then
 If MsgBox(LoadResString_(1300, ""), 4, "", 0, 0) = 6 Then 'É¾³ý²Ù×÷
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1305, "")
  op_remove = True
 End If
ElseIf Button = 1 And op_remove Then
 If MsgBox(LoadResString_(1300, "\\1\\" + Trim(List1.List(List1.ListIndex))), 4, "", 0, 0) = 6 Then  'É¾³ý²Ù×÷
  n% = List1.ListIndex + 1
   Call C_IO.operate_file(1, n%) 'É¾³ý²Ù×÷
 If MsgBox(LoadResString_(1315, ""), 4, "", 0, 0) <> 6 Then 'É¾³ý²Ù×÷
 op_remove = False
 End If
 Else
 If MsgBox(LoadResString_(1315, ""), 4, "", 0, 0) <> 6 Then 'É¾³ý²Ù×÷
 op_remove = False
 End If
 End If
End If
End Sub

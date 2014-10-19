VERSION 5.00
Begin VB.Form ch_ruler 
   Caption         =   "选择进度;点击教科书目录选择学习进度。点击〈选择规则〉按钮,选择推理规则。"
   ClientHeight    =   5085
   ClientLeft      =   765
   ClientTop       =   825
   ClientWidth     =   6690
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5085
   ScaleWidth      =   6690
   Visible         =   0   'False
   Begin VB.CommandButton command4 
      Caption         =   "选择进度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3120
      Width           =   6135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "确定/退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "全部选取"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "选择规则"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "已选择的推理规则："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "可选择的推理规则："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "ch_ruler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim chose_ty As Byte '1选规则 2选进度
Public Sub add_new_item_to_list2(ByVal l%)
Dim i%
If l% = -1 Then
 Exit Sub
End If
For i% = 0 To List2.ListCount - 1
 If List1.ItemData(l%) = List2.ItemData(i%) Then
  Exit Sub
 End If
Next i%
List2.AddItem th_chose(List1.ItemData(l%)).TH_name
th_chose(List1.ItemData(l%)).chose = 1
List2.ItemData(List2.NewIndex) = List1.ItemData(l%)
If l% < List1.ListCount And l% > -1 Then
List1.RemoveItem l%
End If
Label1.Caption = LoadResString_(1260, "\\1\\" + str(List1.ListCount))
Label2.Caption = LoadResString_(1265, "\\1\\" + str(List2.ListCount))
End Sub

Private Sub Command1_Click()
chose_ty = 1
Dim i%
If Command1.Caption = LoadResString_(1270, "") Then
 Command1.Caption = LoadResString_(1275, "")
 Command2.visible = True
 List1.Clear
 List1.width = 2895
 List2.Clear
 List2.visible = True
ch_ruler.Caption = LoadResString_(1275, "")
For i% = -6 To last_th_choose
 If i% <> 0 Then
 If th_chose(i%).chose = 1 And th_chose(i%).chapter <= chapter_no Then
  List2.AddItem th_chose(i%).TH_name
  List2.ItemData(List2.NewIndex) = i%
 ElseIf th_chose(i%).chose = 0 And th_chose(i%).TH_name <> "" And th_chose(i%).chapter <= chapter_no Then
  List1.AddItem th_chose(i%).TH_name
  List1.ItemData(List1.NewIndex) = i%
 ElseIf th_chose(i%).chapter > chapter_no Then
  th_chose(i%).chose = 0
 End If
 End If
Next i%
Label1.Caption = LoadResString_(1260, "\\1\\" + str(List1.ListCount))
Label2.Caption = LoadResString_(1265, "\\1\\" + str(List2.ListCount))
Else
List2.Clear
List1.Clear
For i% = 1 To last_th_choose
th_chose(i%).chose = 0
ch_ruler.List1.AddItem th_chose(i%).TH_name
ch_ruler.List1.ItemData(ch_ruler.List1.NewIndex) = i%
Next i%

Label1.Caption = LoadResString_(1260, "\\1\\" + str(List1.ListCount))
Label2.Caption = LoadResString_(1265, "\\1\\" + str(List2.ListCount))
'add_new_item_to_list2 (List1.ListIndex)
End If
End Sub


Private Sub Command2_Click()
chose_type = 3
Dim i%
'For i% = 1 To last_th_choose
'TH_CHOSE(i%).chose = 1
'Next i%
List1.Clear
List2.Clear
For i% = -6 To last_th_choose
If i% <> 0 Then
If th_chose(i%).chapter <= chapter_no Then
th_chose(i%).chose = 1
ch_ruler.List2.AddItem th_chose(i%).TH_name
ch_ruler.List2.ItemData(ch_ruler.List2.NewIndex) = i%
End If
End If
Next i%
chose_total_theorem = True
Label1.Caption = LoadResString_(1260, "\\1\\" + str(List1.ListCount))
Label2.Caption = LoadResString_(1265, "\\1\\" + str(List2.ListCount))
End Sub

Private Sub Command3_Click()
If chose_type = 1 Then
regist_data.study_progrss = 0
ElseIf chose_type = 2 Then
regist_data.study_progrss = 1
ElseIf chose_type = 3 Then
regist_data.study_progrss = 10000
End If
Dim i%
For i% = -6 To 180
If regist_data.th_chose(i%) <> th_chose(i%).chose And run_type > 2 Then
   MDIForm1.Toolbar1.Buttons(19).visible = True
   MDIForm1.solve.Enabled = True
End If
regist_data.th_chose(i%) = th_chose(i%).chose
Next i%
chose_total_theorem = True
ch_ruler.Hide
End Sub

Private Sub command4_Click()
chose_type = 2
 Command1.Caption = LoadResString_(1270, "")
 Command2.visible = False
ch_ruler.Caption = LoadResString_(1290, "")
List1.Clear
List1.width = 6000
List2.Clear
List2.visible = False
For i% = 1 To last_chapter
 '*** List1.AddItem chapter(i%).text
  List1.ItemData(List1.NewIndex) = i%
Next i%
Label1.Caption = LoadResString_(1295, "")
Label2.Caption = ""
End Sub

Private Sub Form_Load()
Dim i%, k%
ch_ruler.Caption = LoadResString_(4175, "")
Command1.Caption = LoadResString_(4180, "")
Command2.Caption = LoadResString_(4185, "")
command4.Caption = LoadResString_(4190, "")
Command3.Caption = LoadResString_(4195, "")
If regist_data.study_progrss < last_chapter And regist_data.study_progrss > 0 Then
List1.Clear
List1.width = 6000
List2.Clear
List2.visible = False
For i% = 1 To last_chapter
 '*** List1.AddItem chapter(i%).text
  List1.ItemData(List1.NewIndex) = i%
Next i%
Label1.Caption = LoadResString_(1296, "")
Label2.Caption = ""
'***Text1.text = LoadResString_(1280, "") + chapter(regist_data.study_progrss).text
Else
 List1.Clear
 List1.width = 2895
 List2.Clear
 List2.visible = True
 ch_ruler.Caption = LoadResString_(1275, "")
 For i% = -6 To last_th_choose
 If i% <> 0 Then
 If th_chose(i%).chose = 1 Then
  List2.AddItem th_chose(i%).TH_name
  List2.ItemData(List2.NewIndex) = i%
   k% = k% + 1
 ElseIf th_chose(i%).chose = 0 And th_chose(i%).TH_name <> "" Then
  List1.AddItem th_chose(i%).TH_name
  List1.ItemData(List1.NewIndex) = i%
 End If
 End If
Next i%
If k% = 0 Then
   For i% = -6 To last_th_choose
    If i% <> 0 Then
     th_chose(i%).chose = 1
      List2.AddItem th_chose(i%).TH_name
       List2.ItemData(List2.NewIndex) = i%
    End If
   Next i%
End If
End If
End Sub

Private Sub List1_DblClick()
Dim i%, j%
If Command1.Caption = LoadResString_(1275, "") Then
For i% = 0 To List2.ListCount - 1
 If List2.ItemData(i%) > List1.ItemData(List1.ListIndex) Then
  List2.AddItem th_chose(List2.ItemData(List2.ListCount - 1)).TH_name
   List2.ItemData(List2.NewIndex) = List2.ItemData(List2.ListCount - 2)
    For j% = List2.ListCount - 2 To i% + 1 Step -1
     List2.List(j%) = List2.List(j% - 1)
     List2.ItemData(j%) = List2.ItemData(j% - 1)
  Next j%
     List2.List(i%) = List1.List(List1.ListIndex)
     List2.ItemData(i%) = List1.ItemData(List1.ListIndex)
    GoTo ch_ruker_list2out
 End If
Next i%
List2.AddItem th_chose(List1.ItemData(List1.ListIndex)).TH_name
List2.ItemData(List2.NewIndex) = List1.ItemData(List1.ListIndex)
ch_ruker_list2out:
th_chose(List1.ItemData(List1.ListIndex)).chose = 1
If List1.ListIndex < List1.ListCount And List1.ListIndex > -1 Then
List1.RemoveItem List1.ListIndex
End If
Label1.Caption = LoadResString_(1260, "\\1\\" + str(List1.ListCount))
Label2.Caption = LoadResString_(1265, "\\1\\" + str(List2.ListCount))
Else
 '***chapter_no = chapter(List1.ItemData(List1.ListIndex)).no
 regist_data.study_progrss = List1.ItemData(List1.ListIndex)
 If List1.ItemData(List1.ListIndex) < last_chapter Then
 '*** Text1.text = LoadResString_(1280, "") + chapter(List1.ItemData(List1.ListIndex)).text
   For i% = 1 To 160
    If 1 = 1 Then '***th_chose(i%).chapter <= chapter(List1.ItemData(List1.ListIndex)).no Then
     th_chose(i%).chose = 1
    Else
     th_chose(i%).chose = 0
    End If
   Next i%
 Else
 '*** Text1.text = empty_char + chapter(List1.ItemData(List1.ListIndex)).text
 End If
End If
End Sub


Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.Caption = LoadResString_(1275, "") Then

Text1.text = th_chose(List1.ItemData(List1.ListIndex)).text
End If
End Sub

Private Sub List2_DblClick()
Dim i%, j%
If Command1.Caption = LoadResString_(1275, "") Then
For i% = 0 To List1.ListCount - 1
 If List1.ItemData(i%) > List2.ItemData(List2.ListIndex) Then
  List1.AddItem th_chose(List1.ItemData(List1.ListCount - 1)).TH_name
   List1.ItemData(List1.NewIndex) = List1.ItemData(List1.ListCount - 2)
    For j% = List1.ListCount - 2 To i% + 1 Step -1
     List1.List(j%) = List1.List(j% - 1)
     List1.ItemData(j%) = List1.ItemData(j% - 1)
  Next j%
     List1.List(i%) = List2.List(List2.ListIndex)
     List1.ItemData(i%) = List2.ItemData(List2.ListIndex)
    GoTo ch_ruker_list1out
 End If
Next i%
List1.AddItem th_chose(List2.ItemData(List2.ListIndex)).TH_name
List1.ItemData(List1.NewIndex) = List2.ItemData(List2.ListIndex)
ch_ruker_list1out:
th_chose(List2.ItemData(List2.ListIndex)).chose = 0
If List2.ListIndex < List2.ListCount And List2.ListIndex > -1 Then
List2.RemoveItem List2.ListIndex
End If
Label1.Caption = LoadResString_(1260, "\\1\\" + str(List1.ListCount))
Label2.Caption = LoadResString_(1265, "\\1\\" + str(List2.ListCount))
End If
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.Caption = LoadResString_(1275, "") Then
Text1.text = th_chose(List2.ItemData(List2.ListIndex)).text
End If
End Sub

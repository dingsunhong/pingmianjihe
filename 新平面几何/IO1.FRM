VERSION 5.00
Begin VB.Form ioForm 
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   1665
   ClientTop       =   1620
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3975
   ScaleWidth      =   6795
   Begin VB.CommandButton Command2 
      Caption         =   "取　消"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确　认"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   600
      TabIndex        =   3
      Text            =   "*.*"
      Top             =   600
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      TabIndex        =   0
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "盘号"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "目录"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "文件名"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "ioForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim mark As String * 8
Dim i%
Dim path_and_file As String
'Dim of_struct As OFSTRUCT
If Mid$(Dir1.Path, Len(Dir1.Path), 1) = "\" Then
 path_and_file = Dir1.Path & Text1.text
Else
path_and_file = Dir1.Path & "\" & Text1.text
End If
'Call OpenFile(path_and_file, of_struct, 0)
If io_statue = 2 Then
If Text1.text <> "" Then
 Open path_and_file For Random As #2 Len = Len(file_record)
 Call put_wenti_to_record("")
 file_record.record = wenti_record
 file_record.mark = "DOESSOFT"
 Put #2, 1, file_record
Close #2
 End If

Else
If Text1.text <> "" Then
 Open path_and_file For Input As #2

  mark = Input(8, #2)
   If mark <> "DOESSOFT" Then
   Call MsgBox("你打开的不是数学实验室文件，请重试！", 0, "", "", 0)
   Close #2
   Else
   Close #2
    Open path_and_file For Random As #2 Len = Len(file_record)
    Get #2, 1, file_record
   Call init_conditions(0)
'初始化问题条件
wenti_form_title = "答题板—" & Dir1.Path & Text1.text
 Draw_form.Caption = "作图板—" & Dir1.Path & Text1.text
    Wenti_form.Caption = wenti_form_title + "(自动解题)"
wenti_no = 0
old_wenti_no = 0
last_char = 0
Wenti_form.Picture1.Cls
Draw_form.Cls
 
   wenti_record = file_record.record
    Call get_wenti_from_record
    For i% = 0 To wenti_no - 1
    If wenti_cond(i%).no = 0 Then
    Call init_inpcond0(i%)
    End If
    Call cond_to_display(i%)
   Next i%
 Call input_sentences(1, 0, wenti_no - 1, False)
     Call draw_again1
      'Draw_form.Show
      ' open_record = True
Call set_name_for_draw_picture0
For i% = 0 To wenti_no - 1
 Call draw_picture(i%, 0, True)
Next i%
' open_record = False

   End If
  Close #2
End If
End If
ioForm.Hide
End Sub

Private Sub Command2_Click()
Close #2
ioForm.Hide
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Text1.text = File1.FileName
End Sub


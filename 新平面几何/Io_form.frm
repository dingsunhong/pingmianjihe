VERSION 5.00
Begin VB.Form IO_form 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "存例操作"
   ClientHeight    =   3624
   ClientLeft      =   2052
   ClientTop       =   2148
   ClientWidth     =   5244
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3624
   ScaleWidth      =   5244
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   480
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   3120
      Width           =   2292
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
      Height          =   348
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   948
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   2292
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取　消"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确　定"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "插入位置"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   480
      TabIndex        =   7
      Top             =   2760
      Width           =   2772
   End
   Begin VB.Label Label2 
      Caption         =   "选择插入位置"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   2772
   End
   Begin VB.Label Label1 
      Caption         =   "例题名"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   2052
   End
End
Attribute VB_Name = "IO_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim insert_position%
Private Sub Command1_Click()
Dim i%
If io_statue = 1 Then
     If Len(Text1.text) > 0 Then
       Call put_wenti_to_record(Text1.text)
          Call C_IO.operate_file(2, insert_position%) '插入
           Unload IO_form
     Else
       MDIForm1.StatusBar1.Panels(1).text = LoadResString_(5265, "")
     End If
ElseIf io_statue = -1 Then
Get #1, 1, wenti_record
 Call get_wenti_from_record(wenti_record)
  'Close #1
   Unload IO_form
   For i% = 0 To wenti_no - 1
    'Call C_display_wenti.cond_to_display(i%, 1)
   Next i%
 'Call C_display_wenti.input_m_sentences(Wenti_form.Picture1, 1, 1, _
      C_display_wenti.m_last_input_wenti_no, 0, False, 0)
 End If
End Sub

Private Sub Command2_Click()
     'Close #1
      Unload IO_form
End Sub

Private Sub Form_Load()
 Call C_IO.add_exam_name_to_list(Me.List1)
 Me.Text2 = LoadResString_(4155, "")
 insert_position% = C_IO.last_record + 1
 Me.Label1.Caption = LoadResString_(4270, "")
 Me.Label2.Caption = LoadResString_(4275, "")
 Me.Label3.Caption = LoadResString_(4280, "")
 Me.Caption = LoadResString_(4285, "")
 Command1.Caption = LoadResString_(3940, "")
 Command2.Caption = LoadResString_(135, "")
End Sub
Private Sub List1_DblClick()
Me.Text2.text = LoadResString_(4260, "\\1\\" & _
                 Trim(List1.List(List1.ListIndex)))
 insert_position% = List1.ListIndex + 1
End Sub


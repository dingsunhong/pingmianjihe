VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTip 
   BackColor       =   &H00C0FFFF&
   Caption         =   "日积月累"
   ClientHeight    =   5088
   ClientLeft      =   2376
   ClientTop       =   2400
   ClientWidth     =   7728
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5088
   ScaleWidth      =   7728
   StartUpPosition =   2  '屏幕中心
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "Command2"
      Height          =   372
      Index           =   2
      Left            =   6120
      TabIndex        =   12
      Top             =   3960
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command2"
      Height          =   372
      Index           =   1
      Left            =   6120
      TabIndex        =   11
      Top             =   3120
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Index           =   0
      Left            =   6120
      TabIndex        =   10
      Top             =   2280
      Width           =   1332
   End
   Begin VB.PictureBox PictureUp 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   6120
      Picture         =   "frmTip.frx":0000
      ScaleHeight     =   492
      ScaleWidth      =   600
      TabIndex        =   7
      Top             =   480
      Width           =   600
   End
   Begin VB.PictureBox PictureNext 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   6140
      Picture         =   "frmTip.frx":0282
      ScaleHeight     =   492
      ScaleWidth      =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   600
   End
   Begin VB.CheckBox chkLoadTipsAtStartup 
      BackColor       =   &H00C0FFFF&
      Caption         =   "显示这个用法提示小窗口"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   3732
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4764
      ScaleWidth      =   5844
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3492
         Left            =   0
         TabIndex        =   9
         Top             =   1320
         Width           =   5892
         _ExtentX        =   10393
         _ExtentY        =   6160
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmTip.frx":0504
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         Picture         =   "frmTip.frx":05A1
         ScaleHeight     =   852
         ScaleWidth      =   972
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
      Begin VB.PictureBox btrfly2 
         AutoSize        =   -1  'True
         Height          =   972
         Left            =   3600
         Picture         =   "frmTip.frx":2D43
         ScaleHeight     =   943.214
         ScaleMode       =   0  'User
         ScaleWidth      =   943.214
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.PictureBox btrfly1 
         AutoSize        =   -1  'True
         Height          =   972
         Left            =   3600
         Picture         =   "frmTip.frx":39CD
         ScaleHeight     =   943.214
         ScaleMode       =   0  'User
         ScaleWidth      =   943.214
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.PictureBox btrfly 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   924
         Left            =   2400
         Picture         =   "frmTip.frx":4657
         ScaleHeight     =   922.826
         ScaleMode       =   0  'User
         ScaleWidth      =   922.826
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   480
         Width           =   924
         Begin VB.Timer Timer2 
            Interval        =   300
            Left            =   0
            Top             =   0
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "用法提示"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   252
         Left            =   2520
         TabIndex        =   2
         Top             =   120
         Width           =   2172
      End
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The in-memory database of tips.
Dim Tips As New Collection

'No longer needed (from template).
' Name of tips file
'Const TIP_FILE = "TIPOFDAY.TXT"
'Resource file IDs of tips.
Const ID_STR_START = 1145
Const ID_STR_STOP = 1220

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long
Dim flap As Integer

Dim num As Integer
Private Sub butterfly()
    ' Alternate between the two bitmaps.
    If flap = 0 Then
        btrfly.Picture = btrfly1.Picture
        flap = 1
    Else
        btrfly.Picture = btrfly2.Picture
        flap = 0
    End If
End Sub

Private Sub btrfly_Click()

    ''num = num + 1
    If (num Mod 2) Then
    frmTip.Timer2.interval = 0
    Else: frmTip.Timer2.interval = 300
    End If
    num = num + 1
End Sub






Private Sub btrfly1_Click()
    If (num Mod 2) Then
    frmTip.Timer2.interval = 0
    Else: frmTip.Timer2.interval = 300
    End If
    num = num + 1
End Sub

Private Sub btrfly2_Click()
    If (num Mod 2) Then
    frmTip.Timer2.interval = 0
    Else: frmTip.Timer2.interval = 300
    End If
    num = num + 1
End Sub

'Private Sub Command1_Click()
''On Error GoTo command1_click_error
'If index = 1 Then
' Shell ("display.exe"), vbNormalFocus '' vbNormalFocus
' End If
'command1_click_error:
'End Sub

Private Sub PicDisplay_Click()
'On Error GoTo PicDisplay_error
Shell ("Example.exe"), vbNormalFocus
PicDisplay_error:
End Sub


Private Sub PictureDisplay_Click()
'On Error GoTo PictureDisplay_Click_error
Shell ("Display.exe"), vbNormalFocus
PictureDisplay_Click_error:
End Sub


Private Sub Command1_Click(index As Integer)
If index = 0 Then
 MsgBox LoadResString_(260, "") '', vbMsgBoxRtlReading
ElseIf index = 1 Then
'Shell ("操作演示.exe"), vbNormalFocus
'Shell ("Display.exe"), vbNormalFocus
ElseIf index = 2 Then
 Unload Me
End If
End Sub

Private Sub PictureNext_Click()
Dim number1 As Integer
    Dim number2 As Integer
    PictureUp.Enabled = True
    DoNextTip
    num = num + 1
    For number1 = 1 To 1000
    number2 = 16 * number1
    If (num = number2) Then
    frmTip.PictureUp.Enabled = False
    ''Else: frmTip.Commandup.Enabled = True
    End If
    Next number1
End Sub

Private Sub PictureUp_Click()
   DoUpTip
End Sub

Private Sub Timer2_Timer()
    ' Note:  The Interval property of the timer determines
    ' how fast the butterfly's wings flap.
    butterfly
End Sub

Private Sub DoNextTip()
  
    
    ' Select a tip at random.
    'CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' Or, you could cycle through the Tips in order
     
    CurrentTip = CurrentTip + 1
    If Tips.Count < CurrentTip Then
        CurrentTip = 1
    End If
    
    ' Show it.
    frmTip.DisplayCurrentTip

End Sub

Private Sub DoUpTip()

    ' Select a tip at random.
    'CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' Or, you could cycle through the Tips in order

    CurrentTip = CurrentTip - 1
    If (Tips.Count > CurrentTip And CurrentTip = 1 Or Tips.Count < CurrentTip And CurrentTip = 1) Then
        ''CurrentTip = 1
        
        frmTip.PictureUp.Enabled = False
        
    End If
    
    ' Show it.
    frmTip.DisplayCurrentTip

End Sub


'Modified to work with resource string table.
Sub LoadTips()
    Dim NextTip As String
    Dim intString As Integer
    
    ' Each tip read in from resource file.
    For intString = ID_STR_START To ID_STR_STOP
       If intString Mod 5 = 0 Then
        NextTip = LoadResString_(intString, "")
        Tips.Add NextTip
       End If
    Next intString
    Randomize
    ' Display a tip at random.
    DoNextTip

End Sub

Private Sub chkLoadTipsAtStartup_Click()
    SaveSetting App.EXEName, LoadResString_(145, ""), LoadResString_(1750, ""), chkLoadTipsAtStartup.value
End Sub
Private Sub Form_Load()
    Dim ShowAtStartup As Long

    frmTip.PictureUp.Enabled = False
    If CurrentTip = 1 Then ''' CurrentTip = 16
    frmTip.PictureUp.Enabled = False
   End If
    ' See if we should be shown at startup
    ShowAtStartup = GetSetting(App.EXEName, LoadResString_(145, ""), LoadResString_(1750, ""), 1)
     If ShowAtStartup = 0 Then
      ''frmTip.visible = True
     ''' Else
   '''' If ShowAtStartup = 0 Then
        Unload Me
        
        Exit Sub
  
    End If
        
    ' Set the checkbox, this will force the value to be written back out to the registry
    Me.chkLoadTipsAtStartup.value = vbChecked
    
    ' Seed Rnd
    Randomize
   '' Call DoNextTip
    'No longer needed (from template).
    ' Read in the tips file and display a tip at random.
    'If LoadTips(App.Path & "\" & TIP_FILE) = False Then
    '    lblTipText.Caption = "That the " & TIP_FILE & " file was not found? " & vbCrLf & vbCrLf & _
    '       "Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line. " & _
    '       "Then place it in the same directory as the application. "
    'End If

    'Load the tips from the resource file.
    chkLoadTipsAtStartup.Caption = LoadResString_(1750, "")
    Label1.Caption = LoadResString_(1745, "")
    Me.Caption = LoadResString_(1745, "")
    Command1(0).Caption = LoadResString_(515, "")
    Command1(1).Caption = LoadResString_(520, "")
    Command1(2).Caption = LoadResString_(135, "")
    LoadTips
    
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        'richtextbox1=
        RichTextBox1.text = Tips.item(CurrentTip)
    End If
End Sub


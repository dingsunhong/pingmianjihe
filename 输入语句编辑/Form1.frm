VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "转换"
      Height          =   495
      Left            =   4680
      TabIndex        =   12
      Top             =   6840
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   2160
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   5520
      Width           =   11055
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   2160
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   4440
      Width           =   11055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "显示指定输入语句"
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "退出"
      Height          =   495
      Left            =   10080
      TabIndex        =   8
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "修改"
      Height          =   495
      Left            =   7200
      TabIndex        =   7
      Top             =   6840
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2160
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   3480
      Width           =   11055
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   2520
      Width           =   11055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "后一句"
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "前一句"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "启动"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   10680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "语句序号"
      Height          =   495
      Left            =   9360
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub open_data()
Dim last_record As Integer
Dim i%, j%
Open App.Path & "\inpcond.dat" For Random As #1 Len = Len(inpcond0) '
last_record = 0
If LOF(1) > 0 Then '
Do While EOF(1) <> True
Command2_Click
 last_record = last_record + 1
  Get #1, last_record, inpcond0
    inpcond(inpcond0.no).inpcond = inpcond0.inpcond(regist_data.language - 1)
    inpcond(inpcond0.no).ty = inpcond0.ty
    inpcond(inpcond0.no).no = inpcond0.no
    For i% = 0 To 1
    For j% = 0 To 1
    inpcond(inpcond0.no).relation(i%, j%) = inpcond0.relation(i%, j%)
    Next j%
    Next i%
    For i% = 0 To 7
    For j% = 0 To 1
    inpcond(inpcond0.no).taboo(i%).taboo_relation(j%) = inpcond0.taboo(i%).taboo_relation(j%)
    Next j%
    inpcond(inpcond0.no).taboo(i%).ty = inpcond0.taboo(i%).ty
    Next i%
    Loop
Close #1

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Command1_Click()
Open App.Path & "\inpcond_.dat" For Random As #1 Len = Len(inpcond10) '
last_record = 0
Form1.Text1.Text = 0
End Sub

Private Sub Command2_Click()
 last_record = last_record - 1
 Call get_data(last_record)
End Sub

Private Sub Command3_Click()
last_record = last_record + 1
Call get_data(last_record)
End Sub

Private Sub Command4_Click()
Call set_data(last_record)
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Dim i%, wenti_no%
wenti_no% = Text1.Text
For i% = 1 To 150
 Call get_data(i%)
 If inpcond0.no = wenti_no% Then
  Exit Sub
 End If
Next i%
End Sub

Private Sub Command7_Click()
Dim i%
Open App.Path & "\inpcond.dat" For Random As #1 Len = Len(inpcond0) '
Open App.Path & "\inpcond_.dat" For Random As #2 Len = Len(inpcond10) '
last_record = 0
If LOF(1) > 0 Then
Do While EOF(1) <> True
last_record = last_record + 1
Get #1, last_record, inpcond0
inpcond10.no = inpcond0.no
inpcond10.inpcond(0) = inpcond0.inpcond(0)
inpcond10.inpcond(1) = inpcond0.inpcond(1)
inpcond10.inpcond(2) = inpcond0.inpcond(2)
inpcond10.inpcond(3) = inpcond0.inpcond(3)
inpcond10.relation(0, 0) = inpcond0.relation(0, 0)
inpcond10.relation(0, 1) = inpcond0.relation(0, 1)
inpcond10.relation(1, 0) = inpcond0.relation(1, 0)
inpcond10.relation(1, 1) = inpcond0.relation(1, 1)
inpcond10.ty = inpcond0.ty
For i% = 0 To 7
inpcond10.taboo(i%).taboo_relation(0) = inpcond0.taboo(i).taboo_relation(0)
inpcond10.taboo(i%).taboo_relation(1) = inpcond0.taboo(i).taboo_relation(1)
inpcond10.taboo(i%).ty = inpcond0.taboo(i).ty
Next i%
Put #2, last_record, inpcond10
Loop
End If
Close #2
Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #1
End Sub

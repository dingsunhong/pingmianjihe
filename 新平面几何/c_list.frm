VERSION 5.00
Begin VB.Form clist_form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form6"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   240
      Width           =   6135
   End
   Begin VB.CommandButton Confirm 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Exit 
      Caption         =   "�˳�"
      Height          =   615
      Left            =   7200
      TabIndex        =   4
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Three3 
      Caption         =   "������"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Two2 
      Caption         =   "�ڶ���"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton One1 
      Caption         =   "��һ��"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.ListBox List6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   3840
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   6135
   End
End
Attribute VB_Name = "clist_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim volume As Variant


'Private Sub cancel_Click()    ' ��ȡ�������ť�Ĺ��ܴ���
'  clist_form.List6.Clear           ' ��List6�б��ȫ������
'  clist_form.One1.Enabled = True
'  clist_form.Two2.Enabled = True
'  clist_form.Three3.Enabled = True
'End Sub

Private Sub Confirm_Click()    '��ȷ�������ť�Ĺ��ܴ���
 clist_form.List6.Refresh           '   ˢ��List6 �б��ؼ�
 clist_form.One1.Enabled = True     '  ����һ�ᡱ���ť������Ч
 clist_form.Two2.Enabled = True     '  ���ڶ��ᡱ���ť������Ч
 clist_form.Three3.Enabled = True   '  �������ᡱ���ť������Ч
'clist_form.cancel.Enabled = True   '   ��ȡ�������ť������Ч
 clist_form.Exit.Enabled = True     '   ���˳������ť������Ч
End Sub

Private Sub exit_Click()         '���˳������ť���ܴ���
  End
End Sub

Private Sub One1_Click()       '����һ�ᡱ���ť�Ĺ��ܴ���
clist_form.List6.Clear              ' ��List6�б��ȫ������
clist_form.One1.Enabled = True      '����һ�ᡱ���ť������Ч
clist_form.Two2.Enabled = True      ' ���ڶ��ᡱ���ť������Ч
clist_form.Three3.Enabled = True    '  �������ᡱ���ť������Ч
'clist_form.cancel.Enabled = True   '   ��ȡ�������ť������Ч
clist_form.Exit.Enabled = True      '   ���˳������ť������Ч
clist_form.Confirm.Enabled = True   '  ��ȷ�� �����ť������Ч
 Call set_gate1(1)             '������ʾ��һ��Ŀ¼�Ĺ��̺�����set_gate1(1)��
End Sub


Private Sub Text6_Click()      '��List6���ı���Ĺ��ܴ���
clist_form.Text6.Refresh            '�ԡ�List6���ı���ؼ�ˢ��
clist_form.Text6.SetFocus
'clist_form.Text6.Visible
End Sub                        '��ı���ľ��幦����Ҫ������

Private Sub Three3_Click()     '�������ᡱ���ť�Ĺ��ܴ���
clist_form.List6.Clear              ' ��List6�б��ȫ������
clist_form.One1.Enabled = True      '  ����һ�ᡱ���ť������Ч
clist_form.Two2.Enabled = True      ' ���ڶ��ᡱ���ť������Ч
clist_form.Three3.Enabled = True    '  �������ᡱ���ť������Ч
'clist_form.cancel.Enabled = True   '   ��ȡ�������ť������Ч
clist_form.Exit.Enabled = True      '   ���˳������ť������Ч
clist_form.Confirm.Enabled = True   '  ��ȷ�� �����ť������Ч
 Call set_gate3(3)             '������ʾ������Ŀ¼�Ĺ��̺�����set_gate3(3)��
End Sub

Private Sub Two2_Click()       '���ڶ��ᡱ���ť�Ĺ��ܴ���
clist_form.List6.Clear              ' ��List6�б��ȫ������
clist_form.Two2.Enabled = True      ' ���ڶ��ᡱ���ť������Ч
clist_form.One1.Enabled = True      '  ����һ�ᡱ���ť������Ч
clist_form.Three3.Enabled = True    '  ����һ�ᡱ���ť������Ч
'clist_form.cancel.Enabled = True   '��ȡ�������ť������Ч
clist_form.Exit.Enabled = True           ' ���˳������ť������Ч
clist_form.Confirm.Enabled = True   '   ��ȷ�������ť������Ч
 Call set_gate2(2)             '������ʾ�ڶ���Ŀ¼�Ĺ��̺�����set_gate2(2)��
 
End Sub

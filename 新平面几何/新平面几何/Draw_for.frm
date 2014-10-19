VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Draw_form 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "×÷Í¼°å"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   1740
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   15  'Merge Pen Not
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   430
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Visible         =   0   'False
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   852
      Left            =   960
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   750
      Left            =   2280
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   360
      Top             =   360
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6855
      Left            =   9360
      Max             =   200
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      Left            =   6960
      Max             =   480
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin ComctlLib.ImageList ImageList4 
      Left            =   1680
      Top             =   4200
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   26
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":0292
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":0524
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":07B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":0A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":0F6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":11FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":1490
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":1722
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":19B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":1C46
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":1ED8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":216A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":23FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":268E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":2920
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":2BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":2E44
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":30D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":3368
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":35FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":388C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":3B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":3DB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":4042
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   3240
      Top             =   2880
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   26
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":42D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":43E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":44F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":460A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":471C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":482E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":4940
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":4A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":4B64
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":4C76
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":4D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":4E9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":4FAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":50BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":51D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":52E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":53F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":5506
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":5618
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":572A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":583C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":594E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":5A60
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":5BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":5CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":5DD6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList3 
      Left            =   2400
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   30
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":5EE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":617A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":640C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":669E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":6930
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":6BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":6E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":70E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":7378
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":760A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":789C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":7B2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":7DC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":8052
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":82E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":8576
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":8808
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":8A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":8D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":8FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":9250
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":94E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":9774
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":9A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":9C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":9F2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":A1BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":A44E
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":A6E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":A972
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   30
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":AC04
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":AD26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":AE48
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":AF6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":B08C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":B1AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":B2D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":B3F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":B514
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":B636
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":B758
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":B87A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":B99C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":BABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":BBE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":BD02
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":BE24
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":BF46
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":C068
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":C18A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":C2AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":C3CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":C4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":C612
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":C734
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":C856
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":C978
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":CA9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":CBBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Draw_for.frx":CCDE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   2085
   End
End
Attribute VB_Name = "Draw_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Option Explicit
Dim tangent_point_type%
Dim view_p As POINTAPI
Dim t_vscroll_value As Integer
'Dim remove_p   As String
Dim change_operat_statue As Boolean
Dim change_op_info As String
Dim temp_poi As point_type
Dim text_type As Byte
Dim text_text As String
Dim measur_step As Byte
Dim tk1 As Single
Dim tk2 As Single
Dim button_statue As Single 'Êó±ê°´¼ü×´Ì¬
Dim display_operat_statue  As String
Dim operator_statue As Integer
Private Sub Form_Activate()
Dim i%
If event_statue = wait_for_draw_picture Then
For i% = C_display_wenti.m_last_input_wenti_no To C_display_wenti.m_last_input_wenti_no - 1
Call draw_picture(i%, 0, False)
Next i%
End If
If arrange_window_type = 2 Then
  'arrange_window_type = 1
Wenti_form.left = 0
Wenti_form.top = 0
Wenti_form.width = Screen.width - 280
Wenti_form.Height = Screen.Height - 1550 + int_w_y
Draw_form.width = Screen.width - 280
Draw_form.left = 100
Draw_form.top = 320
Draw_form.Height = Screen.Height - 1550 + int_w_y
 Draw_form.Picture1.Height = Draw_form.ScaleHeight
  Draw_form.Picture1.width = Draw_form.ScaleWidth
 Draw_form.Refresh
   Wenti_form.Refresh
    Draw_form.SetFocus
End If
End Sub

Private Sub Form_DblClick()
If yidian_type = 19 Then
Call C_curve.move_pucture_along_curve
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Call plane_geometry_key_press(KeyAscii)
End Sub

Private Sub Form_Load()
Dim i As Integer
'ÉèÖÃ»­Í¼°å´óÐ¡
Draw_form.Height = screeny& - 1350
Draw_form.width = (screenx& - 250) / 2 - 10
'ÉèÖÃ»­Í¼°åµÄ¹ö¶¯
 Draw_form.Picture1.Height = Draw_form.ScaleHeight
  Draw_form.Picture1.width = Draw_form.ScaleWidth
'Call draw_ruler(ratio_for_measure.ratio_for_measure, display)
'remove_p = LoadResString_(114)
'¡®Call set_pen
choose_point = -1
'drag_statue = 0
Ratio_for_measure.Ratio_for_measure = 0

'For i = 0 To 50
'poi(i).data(0).data0.color = 9
'poi(i).name = ""
'Next i
'For i = 0 To 30
'lin(i).data(0).data0.color = 0
'Next i
'For i = 0 To 20
'circ(i).data(0).data0.color = 0
'Next i
'centerdisplay = 0
'Picture1.AutoRedraw = 1
'extense_no = 0

   
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If protect_munu = 0 Then
 Call plane_geometry_draw_mouse_down(Button, Shift, X, Y)
'Else
' protect_munu = 0
' protect_munu_ = 0
'End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call plane_geometry_draw_mouse_move(Button, Shift, X, Y)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call plane_geometry_draw_mouse_up(Button, Shift, X, Y)
End Sub


Private Sub Form_Resize()
If arrange_window_type = 0 Then
Wenti_form.left = Draw_form.width + 5
If screenx& > Draw_form.width Then
Wenti_form.width = Screen.width - Draw_form.width
End If
Draw_form.Height = Screen.Height - 1350 + int_w_y
Draw_form.top = 0
Draw_form.left = 0
ElseIf arrange_window_type = 1 Then
Wenti_form.width = Screen.width - 150
Draw_form.width = Screen.width - 150
Draw_form.top = 0
Draw_form.left = 0
Wenti_form.left = 0
Wenti_form.top = Draw_form.Height + 5
If Screen.Height - Wenti_form.top - 1350 + int_w_y > 0 Then
Wenti_form.Height = Screen.Height - Wenti_form.top - 1350 + int_w_y
End If
Draw_form.VScroll1.left = Draw_form.ScaleWidth - 16
Draw_form.VScroll1.Height = Draw_form.ScaleHeight
 Draw_form.Picture1.Height = Draw_form.ScaleHeight
  Draw_form.Picture1.width = Draw_form.ScaleWidth
End If
If Draw_form.AutoRedraw = False Then
 Call BitBlt(Draw_form.hdc, 0, 0, Draw_form.Picture1.width, _
     Draw_form.Picture1.Height, Draw_form.Picture1.hdc, 0, 0, &H8800C6)
End If
End Sub

Private Sub HScroll1_Change()
Dim m%, i%
Dim n As Byte
If HScroll1.value > 1 And HScroll1.value < 480 Then
m% = HScroll1.value
If m% <> Ratio_for_measure.Ratio_for_measure Then
 Call draw_ruler(Ratio_for_measure.Ratio_for_measure, delete)
  Ratio_for_measure.Ratio_for_measure = m%
   Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
    Call measur_again
End If
End If
End Sub



Private Sub list1_Click()
Dim respon As Boolean
 ' list1.List (list1.ListIndex)
 List1.visible = False
If operator = "draw_point_and_line" Then
 list_type_for_draw = List1.ListIndex + 1
 If List1.ListIndex = 0 Then
 MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2180, "")
 ElseIf List1.ListIndex = 1 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2185, "")
 ElseIf List1.ListIndex = 2 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2190, "")
 ElseIf List1.ListIndex = 3 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2195, "")
 ElseIf List1.ListIndex = 4 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2200, "")
 End If
ElseIf operator = "draw_circle" Then
 list_type_for_draw = List1.ListIndex + 1
  If list_type_for_draw = 1 Then
   MDIForm1.StatusBar1.Panels(1).text = _
      LoadResString_(2205, "\\1\\" + m_poi(temp_point(0).no).data(0).data0.name)
  ElseIf list_type_for_draw = 2 Then
   MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2210, "")
  ElseIf list_type_for_draw = 3 Then
   MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2215, "")
  End If

ElseIf operator = "move_point" Then
 list_type_for_draw = List1.ListIndex + 1

 If List1.ListIndex = 0 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2220, "")
 ElseIf List1.ListIndex = 1 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2225, "")
 End If
  List1.visible = False
   HScroll1.visible = False
ElseIf operator = "epolygon" Then
 list_type_for_draw = List1.ListIndex + 1
  If List1.ListIndex = 0 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2000, "")
 ElseIf List1.ListIndex = 1 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2010, "")
 ElseIf List1.ListIndex = 2 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2015, "")
 ElseIf List1.ListIndex = 3 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2020, "")
 End If

ElseIf operator = "measure" Then
 list_type_for_draw = List1.ListIndex + 1
 If List1.ListIndex = 0 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2230, "")
 ElseIf List1.ListIndex = 1 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2235, "")
 ElseIf List1.ListIndex = 2 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2240, "")
 ElseIf List1.ListIndex = 3 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2245, "")
 End If
ElseIf operator = "re_name" Then
 If list_type_for_draw = 3 Then
 If List1.ListIndex = 0 Then
  If temp_point(0).no = last_conditions.last_cond(1).point_no Then
   If MsgBox(LoadResString_(1805, "\\1\\" + m_poi(temp_point(0).no).data(0).data0.name), vbYesNo, "", "", 0) = 6 Then
   Else
  End If
  Else
  Call MsgBox(LoadResString_(1790, "\\1\\" + m_poi(temp_point(0).no).data(0).data0.name + _
                       "\\2\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name), 0, "", "", 0)
'       Call C_display_picture.redraw_point(temp_point(0))  ', display)
  End If
 ElseIf List1.ListIndex = 1 Then
 '  Call C_display_picture.redraw_point(temp_point(0))  ', display)
  'poi(temp_point(0)).data(0).data0.color = 9
 End If
List1.visible = False

operat_is_acting = False
End If
ElseIf operator = "change_picture" Then
If draw_step <> 1 Then
 list_type_for_draw = List1.ListIndex + 1
' Else
 If list_type_for_draw = 1 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1810, "")
 ElseIf list_type_for_draw = 2 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1815, "")
 ElseIf list_type_for_draw = 3 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1820, "")
 ElseIf list_type_for_draw = 4 Then
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1825, "")
 End If
End If
ElseIf operator = "set" Then
If List1.ListIndex = 0 Then
 HScroll1.visible = True
  HScroll1.value = Ratio_for_measure.Ratio_for_measure
    MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1830, "")
 
Else
  HScroll1.visible = False
End If
End If

End Sub






Public Function find_last_input0(ByVal n%) As Integer
Dim i%, j%
For i% = 1 To C_display_wenti.m_last_input_wenti_no
If C_display_wenti.m_no(i%) = 0 Then
 n% = 6
  For j% = 0 To 5
   If C_display_wenti.m_condition(i%, j%) = empty_char Then
       n% = j%
       find_last_input0 = i%
        Exit Function
  End If
 Next j%
End If
Next i%


End Function



Public Sub round_change(ByVal c_x&, ByVal c_y&, _
   ByVal s!, ByVal c!, ByVal X&, ByVal Y&, out_x&, out_y&)
out_x& = c_x& + c! * (X& - c_x&) - s! * (Y& - c_y&)
out_y& = c_x& + c! * (Y& - c_y&) - s! * (X& - c_x&)
 
End Sub




Public Function can_not_op_action(ByVal p_ty%, ByVal ele1%, ByVal ele2%, _
     ByVal d_or_up As Boolean) As Boolean
Select Case operator

 Case "draw_point_and_line"
  If list_type_for_draw = 4 Then
   If draw_step = 1 Then
    If p_ty% = new_point_on_circle Then
     If ele1% <> temp_circle(0) Then
      can_not_op_action = True
     End If
    ElseIf p_ty% = new_point_on_line_circle Then
     If ele2% <> temp_circle(0) Then
      can_not_op_action = True
     End If
    ElseIf p_ty% = new_point_on_circle_circle Then
     If ele1% <> temp_circle(0) And ele2% <> temp_circle(0) Then
      can_not_op_action = True
     End If
    Else
      can_not_op_action = True
    End If
   End If
  ElseIf list_type_for_draw = 5 Then
    If (draw_step = 4 And d_or_up = False) Or _
         (draw_step = 5 And d_or_up) Then
    If p_ty% = new_point_on_circle Then
     If ele1% <> temp_circle(0) Then
      can_not_op_action = True
     End If
    ElseIf p_ty% = new_point_on_line_circle Then
     If ele2% <> temp_circle(0) Then
      can_not_op_action = True
     End If
    ElseIf p_ty% = new_point_on_circle_circle Or _
         p_ty% = new_point_on_circle_circle12 Or _
           p_ty% = new_point_on_circle_circle21 Then
     If ele1% <> temp_circle(0) And ele2% <> temp_circle(0) Then
      can_not_op_action = True
     End If
    Else
      can_not_op_action = True
    End If
   End If
  End If
Case "draw_circle"
  If list_type_for_draw = 1 Then
  ElseIf list_type_for_draw = 2 Then
  ElseIf list_type_for_draw = 3 Then
  ElseIf list_type_for_draw = 4 Then
  End If
Case "paral_verti"
  If list_type_for_draw = 1 Then
  ElseIf list_type_for_draw = 2 Then
  ElseIf list_type_for_draw = 3 Then
  ElseIf list_type_for_draw = 4 Then
  End If
 Case "epolygon"
  If list_type_for_draw = 1 Then
  ElseIf list_type_for_draw = 2 Then
  ElseIf list_type_for_draw = 3 Then
  ElseIf list_type_for_draw = 4 Then
  End If
Case "move_point"
  If list_type_for_draw = 1 Then
  ElseIf list_type_for_draw = 2 Then
  ElseIf list_type_for_draw = 3 Then
  ElseIf list_type_for_draw = 4 Then
  End If
Case "change_picture"
  If list_type_for_draw = 1 Then
  ElseIf list_type_for_draw = 2 Then
  ElseIf list_type_for_draw = 3 Then
  ElseIf list_type_for_draw = 4 Then
  End If
Case "measure"
  If list_type_for_draw = 1 Then
  ElseIf list_type_for_draw = 2 Then
  ElseIf list_type_for_draw = 3 Then
  ElseIf list_type_for_draw = 4 Then
  End If
Case "set"
  If list_type_for_draw = 1 Then
  ElseIf list_type_for_draw = 2 Then
  ElseIf list_type_for_draw = 3 Then
  ElseIf list_type_for_draw = 4 Then
  End If
End Select

End Function

Private Sub Picture1_Resize()
'Call BitBlt(Draw_form.Picture1.hdc, 0, 0, Draw_form.Picture1.Width, _
     Draw_form.Picture1.Height, Draw_form.hdc, 0, 0, &H8800C6)
End Sub
Private Sub Timer1_Timer()
draw_time_act = True
End Sub

'Private Sub Timer1_Timer()
'If re_name_ty = 1 Then
're_name_ty = 0
'choose_point = 1
' Call C_display_picture.draw_red_point(choose_point)
'yidian_stop = False
'Draw_form.SetFocus
'Call C_display_picture.flash_point(choose_point)
'End If
'End Sub

Private Sub VScroll1_Change()
Dim i%
Dim m_coord As POINTAPI
'Call C_display_picture.Backup_picture
m_coord.Y = t_vscroll_value - VScroll1.value
For i% = 0 To last_conditions.last_cond(1).point_no
t_coord.X = m_coord.X
t_coord.Y = m_poi(i%).data(0).data0.coordinate.Y + m_coord.Y
Call set_point_coordinate(i%, t_coord, True)
' Call set_point_color(i%, 15)
Next i%
Call change_picture_(0, 0)
t_vscroll_value = VScroll1.value
End Sub

Public Sub set_caption()

End Sub

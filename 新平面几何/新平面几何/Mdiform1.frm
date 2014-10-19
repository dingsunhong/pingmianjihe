VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "DDS-平面几何"
   ClientHeight    =   6435
   ClientLeft      =   90
   ClientTop       =   525
   ClientWidth     =   9420
   Icon            =   "Mdiform1.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "MDIForm1"
   MousePointer    =   1  'Arrow
   Picture         =   "Mdiform1.frx":0442
   ScrollBars      =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   21
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "new"
            Object.ToolTipText     =   "新问题"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "open"
            Object.ToolTipText     =   "打开文件"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "save"
            Object.ToolTipText     =   "存盘"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "print"
            Object.ToolTipText     =   "打印"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "point"
            Object.ToolTipText     =   "画点和线"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "circle"
            Object.ToolTipText     =   "画圆"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "paral"
            Object.ToolTipText     =   "平行垂直"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "poly"
            Object.ToolTipText     =   "画多边形"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "anima"
            Object.ToolTipText     =   "画动态图"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "change"
            Object.ToolTipText     =   "图形变换"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "measur"
            Object.ToolTipText     =   "测量"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "set"
            Object.ToolTipText     =   "设置标尺"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ask"
            Object.ToolTipText     =   "提问"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "name"
            Object.ToolTipText     =   "修改输入"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "postpone"
            Object.ToolTipText     =   "暂停控制"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "stop"
            Object.ToolTipText     =   "终止解题"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "run"
            Object.ToolTipText     =   "解题"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "markpen"
            Object.ToolTipText     =   "标示笔"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "un_do"
            Object.ToolTipText     =   "撤消最后一步操作"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7320
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5640
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   3840
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1320
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2160
      Top             =   1200
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      DragMode        =   1  'Automatic
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6135
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   14113
            MinWidth        =   14113
            Picture         =   "Mdiform1.frx":0888
            Object.Tag             =   ""
            Object.ToolTipText     =   "操作提示和信息显示"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   1677
            MinWidth        =   1677
            TextSave        =   "2014-2-11"
            Object.Tag             =   ""
            Object.ToolTipText     =   "日期"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "0:39"
            Object.Tag             =   ""
            Object.ToolTipText     =   "时间"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   34
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":0A62
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":0B74
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":0C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":0D98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":0EAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":0FBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":11E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":12F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":1404
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":1516
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":1628
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":173A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":184C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":195E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":1A70
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":1B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":1C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":1DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":1EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":1FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":20DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":21EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":2300
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":2412
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":2524
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":2636
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":2748
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":285A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":296C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":2A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":2B90
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":2CA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mdiform1.frx":2DB4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "文件[&F]"
      Begin VB.Menu new 
         Caption         =   "新  建[&N] "
         HelpContextID   =   57600
      End
      Begin VB.Menu Open 
         Caption         =   "打  开[&O]"
         HelpContextID   =   57601
      End
      Begin VB.Menu save 
         Caption         =   "存  盘[&S]"
         HelpContextID   =   57602
      End
      Begin VB.Menu save_as 
         Caption         =   "另存为[&A]"
         HelpContextID   =   57603
      End
      Begin VB.Menu savee 
         Caption         =   "存  例[&E]　"
         HelpContextID   =   57604
      End
      Begin VB.Menu mprint 
         Caption         =   "打  印[&P]"
         HelpContextID   =   57605
      End
      Begin VB.Menu Exit 
         Caption         =   "退  出[&X]"
         HelpContextID   =   57606
      End
   End
   Begin VB.Menu edit 
      Caption         =   "编辑[&E]"
      Begin VB.Menu set_mode 
         Caption         =   "设置"
      End
      Begin VB.Menu re_name 
         Caption         =   "修改输入"
         Begin VB.Menu re_name_all 
            Caption         =   "所有点重命名"
            HelpContextID   =   57613
         End
         Begin VB.Menu re_name_one 
            Caption         =   "修改点的属性"
            HelpContextID   =   57614
         End
         Begin VB.Menu remove_point_ 
            Caption         =   "删除点"
            HelpContextID   =   57615
         End
         Begin VB.Menu note 
            Caption         =   "注释"
            Visible         =   0   'False
         End
         Begin VB.Menu delete_last_op 
            Caption         =   "撤消最后一步操作"
            HelpContextID   =   57619
         End
      End
      Begin VB.Menu set_p 
         Caption         =   "设置作图板属性"
         HelpContextID   =   57616
      End
   End
   Begin VB.Menu draw 
      Caption         =   "作图[&D]"
      Begin VB.Menu porine 
         Caption         =   "画点或线"
         Begin VB.Menu pandline 
            Caption         =   "画点和线"
            HelpContextID   =   57510
         End
         Begin VB.Menu midpoint 
            Caption         =   "作线段中点"
            HelpContextID   =   57511
         End
         Begin VB.Menu ratio_point 
            Caption         =   "作线段的定比分点"
            HelpContextID   =   57512
         End
         Begin VB.Menu line_given_length 
            Caption         =   "作定长线段"
            HelpContextID   =   57513
         End
         Begin VB.Menu equal_line 
            Caption         =   "作等长线段"
            HelpContextID   =   57514
         End
         Begin VB.Menu equal_angle_line 
            Caption         =   "作角平分线"
            HelpContextID   =   57515
         End
      End
      Begin VB.Menu draw_circle 
         Caption         =   "画圆"
         Begin VB.Menu d_circle 
            Caption         =   "画圆"
            HelpContextID   =   57520
         End
         Begin VB.Menu tangent_line_point_to_circle 
            Caption         =   "画点到圆的切线"
            HelpContextID   =   57522
         End
         Begin VB.Menu tangent_of_two_circles 
            Caption         =   "画两圆公切线"
         End
         Begin VB.Menu draw_circle_tnagent_to_circle 
            Caption         =   "画一圆相切于已有圆"
         End
         Begin VB.Menu draw_circle_tngent_to_two_circles 
            Caption         =   "画两圆的公切圆"
         End
      End
      Begin VB.Menu paralandverti0 
         Caption         =   "作平行垂直线"
         Begin VB.Menu paralandverti 
            Caption         =   "作平行垂直线"
            HelpContextID   =   57530
         End
         Begin VB.Menu verti_mid_line 
            Caption         =   "作垂直平分线"
            HelpContextID   =   57531
         End
      End
      Begin VB.Menu E_polygon 
         Caption         =   "作正多边形"
         Begin VB.Menu e_polygon3 
            Caption         =   "正三角形"
            HelpContextID   =   57540
         End
         Begin VB.Menu e_polygon4 
            Caption         =   "正四边形"
            HelpContextID   =   57541
         End
         Begin VB.Menu e_polygon5 
            Caption         =   "正五边形"
            HelpContextID   =   57542
         End
         Begin VB.Menu e_polygon6 
            Caption         =   "正六边形"
            HelpContextID   =   57543
         End
      End
      Begin VB.Menu move_picture 
         Caption         =   "作动态图"
         HelpContextID   =   57550
         Begin VB.Menu set_view_point 
            Caption         =   "设置观察点"
         End
         Begin VB.Menu set_opera_type 
            Caption         =   "设置操作方式"
            Begin VB.Menu by_hand 
               Caption         =   "手动"
            End
            Begin VB.Menu auto 
               Caption         =   "自动"
            End
         End
      End
      Begin VB.Menu change_picture 
         Caption         =   "图形变换"
         HelpContextID   =   57560
         Begin VB.Menu set_picture_for_change 
            Caption         =   "设置变换图形"
            Begin VB.Menu set_change_line 
               Caption         =   "设置变换直线"
            End
            Begin VB.Menu set_polygon_for_change 
               Caption         =   "设置变换多边形"
            End
            Begin VB.Menu set_circle_for_change 
               Caption         =   "设置变换圆"
            End
         End
         Begin VB.Menu set_change_type 
            Caption         =   "设置变换方式"
            Enabled         =   0   'False
            Begin VB.Menu move_part 
               Caption         =   "平移"
            End
            Begin VB.Menu turn_part 
               Caption         =   "旋转"
            End
            Begin VB.Menu turn_over 
               Caption         =   "翻转"
            End
            Begin VB.Menu fd_sx_p 
               Caption         =   "放大缩小"
            End
            Begin VB.Menu zhouduicheng 
               Caption         =   "轴对称图"
            End
            Begin VB.Menu zhongxinduichen 
               Caption         =   "中心对称"
            End
            Begin VB.Menu initial_picture 
               Caption         =   "恢复原图"
               Visible         =   0   'False
            End
            Begin VB.Menu delete_picture_change 
               Caption         =   "撤消图形变换"
            End
         End
      End
   End
   Begin VB.Menu judge 
      Caption         =   "判断"
      Visible         =   0   'False
      Begin VB.Menu chose1 
         Caption         =   "选择一"
      End
      Begin VB.Menu chose2 
         Caption         =   "选择二"
      End
      Begin VB.Menu chose3 
         Caption         =   "选择三"
      End
   End
   Begin VB.Menu Inputcond 
      Caption         =   "条件[&T]"
      HelpContextID   =   57630
      Begin VB.Menu s1 
         Caption         =   "一般条件"
         Begin VB.Menu S12 
            Caption         =   "□□＝□□"
         End
         Begin VB.Menu S14 
            Caption         =   "∠□□□=∠□□□"
         End
         Begin VB.Menu S15 
            Caption         =   "∠□□□=_"
         End
         Begin VB.Menu s16 
            Caption         =   "□□＝_"
         End
         Begin VB.Menu S17 
            Caption         =   "□□:□□＝_"
         End
         Begin VB.Menu s18 
            Caption         =   "□□+□□＝_"
         End
         Begin VB.Menu s19 
            Caption         =   "□□：□□＝□□：□□"
         End
         Begin VB.Menu s1A 
            Caption         =   "□□＝□□+□□"
         End
         Begin VB.Menu s1B 
            Caption         =   "∠□□□:∠□□□=_"
         End
         Begin VB.Menu s1D 
            Caption         =   "□□是∠□□□的平分线"
         End
         Begin VB.Menu s1C 
            Caption         =   "线段的一般关系式"
         End
         Begin VB.Menu s1G 
            Caption         =   "_、_是一元二次方程_的两个根"
         End
      End
      Begin VB.Menu S2 
         Caption         =   "直线和直线上的点"
         Begin VB.Menu S21 
            Caption         =   "直线□□上任取一点□"
         End
         Begin VB.Menu S22 
            Caption         =   "□□∥□□"
         End
         Begin VB.Menu S23 
            Caption         =   "□□⊥□□"
         End
         Begin VB.Menu S24 
            Caption         =   "在□□的垂直平分线上取任一点□"
         End
         Begin VB.Menu S31 
            Caption         =   "取线段□□的中点□"
         End
         Begin VB.Menu S32 
            Caption         =   "□是线段□□上分比为_的分点"
         End
         Begin VB.Menu S34 
            Caption         =   "过□作直线□□的垂线垂足为□"
         End
      End
      Begin VB.Menu S4 
         Caption         =   "圆和关于圆的点线"
         Begin VB.Menu S41 
            Caption         =   "⊙□_上任取一点□"
         End
         Begin VB.Menu S42 
            Caption         =   "过点□、□、□作⊙□"
         End
         Begin VB.Menu S4C 
            Caption         =   "⊙□_和⊙□_相切于点□ "
         End
         Begin VB.Menu S45 
            Caption         =   "过□作⊙□_的切线□□"
         End
         Begin VB.Menu S4A 
            Caption         =   "弧□□＝弧□□"
         End
      End
      Begin VB.Menu S5 
         Caption         =   "△的心"
         Begin VB.Menu S51 
            Caption         =   "□是△□□□的重心"
         End
         Begin VB.Menu S52 
            Caption         =   "□是△□□□的外接圆的圆心"
         End
         Begin VB.Menu S53 
            Caption         =   "□是△□□□的垂心"
         End
         Begin VB.Menu S54 
            Caption         =   "□是△□□□的内切圆的圆心"
         End
      End
      Begin VB.Menu S6 
         Caption         =   "多边形"
         Begin VB.Menu S61 
            Caption         =   "任意△□□□"
         End
         Begin VB.Menu S6E 
            Caption         =   "△□□□的周长="
         End
         Begin VB.Menu S6F 
            Caption         =   "△□□□的面积="
         End
         Begin VB.Menu S6I 
            Caption         =   "△□□□≌△□□□"
         End
         Begin VB.Menu S6J 
            Caption         =   "△□□□∽△□□□"
         End
         Begin VB.Menu S62 
            Caption         =   "任意四边形□□□□"
         End
         Begin VB.Menu S6G 
            Caption         =   "四边形□□□□的周长="
         End
         Begin VB.Menu S6H 
            Caption         =   "四边形□□□□的面积="
         End
         Begin VB.Menu S63 
            Caption         =   "△□□□是等腰三角形"
         End
         Begin VB.Menu S64 
            Caption         =   "△□□□是等腰直角三角形"
         End
         Begin VB.Menu S65 
            Caption         =   "△□□□是等边三角形"
         End
         Begin VB.Menu S66 
            Caption         =   "□□□□是梯形"
         End
         Begin VB.Menu S67 
            Caption         =   "□□□□是等腰梯形"
         End
         Begin VB.Menu S68 
            Caption         =   "□□□□是长方形"
         End
         Begin VB.Menu S69 
            Caption         =   "□□□□是正方形"
         End
         Begin VB.Menu S6A 
            Caption         =   "□□□□是平行四边形"
         End
         Begin VB.Menu S6B 
            Caption         =   "□□□□是菱形"
         End
         Begin VB.Menu S6C 
            Caption         =   "□□□□□是正五边形"
         End
         Begin VB.Menu S6D 
            Caption         =   "□□□□□□是正六边形"
         End
      End
   End
   Begin VB.Menu conclusion 
      Caption         =   "结论[&C]"
      HelpContextID   =   57640
      Begin VB.Menu C_line 
         Caption         =   "关于直线的结论"
         Begin VB.Menu c_line1 
            Caption         =   "□、□、□三点共线"
         End
         Begin VB.Menu c_line2 
            Caption         =   "□□、□□、□□三线共点"
         End
         Begin VB.Menu c_line3 
            Caption         =   "□□平行于□□"
         End
         Begin VB.Menu c_line4 
            Caption         =   "□□垂直于□□"
         End
         Begin VB.Menu c_line5 
            Caption         =   "点□位于线段□□的垂直平分线上"
         End
      End
      Begin VB.Menu cl1 
         Caption         =   "关于线段的结论"
         Begin VB.Menu c_linel1_0 
            Caption         =   "□□＝_"
         End
         Begin VB.Menu c_linel1_1 
            Caption         =   "点□是线段□□的中点"
         End
         Begin VB.Menu c_line1_2 
            Caption         =   "□□：□□＝_"
         End
         Begin VB.Menu c_line1_3 
            Caption         =   "□□＝□□"
         End
         Begin VB.Menu c_line1_4 
            Caption         =   "□□：□□＝□□：□□"
         End
         Begin VB.Menu c_line1_5 
            Caption         =   "其他线段表达式"
         End
         Begin VB.Menu c_line1_6 
            Caption         =   "_=定值"
         End
      End
      Begin VB.Menu c_angle 
         Caption         =   "关于角的结论"
         Begin VB.Menu c_angle0 
            Caption         =   "∠□□□=_"
         End
         Begin VB.Menu c_angle1 
            Caption         =   "∠□□□=∠□□□"
         End
         Begin VB.Menu c_angle2 
            Caption         =   "∠□□□:∠□□□=_"
         End
         Begin VB.Menu c_angle3 
            Caption         =   "∠□□□+∠□□□=_"
         End
         Begin VB.Menu c_angle4 
            Caption         =   "∠□□□=∠□□□+∠□□□"
         End
      End
      Begin VB.Menu c_triangle 
         Caption         =   "关于三角形的结论"
         Begin VB.Menu c_triangle1 
            Caption         =   "△□□□≌△□□□"
         End
         Begin VB.Menu c_triangle2 
            Caption         =   "△□□□∽△□□□"
         End
         Begin VB.Menu c_triangle8 
            Caption         =   "△□□□与△□□□面积相等"
         End
         Begin VB.Menu c_triangle3 
            Caption         =   "△□□□是等边三角形"
         End
         Begin VB.Menu c_triangle5 
            Caption         =   "△□□□是等腰三角形"
         End
         Begin VB.Menu c_triangle4 
            Caption         =   "△□□□是等腰直角三角形"
         End
         Begin VB.Menu c_triangle6 
            Caption         =   "△□□□的周长="
         End
         Begin VB.Menu c_triangle7 
            Caption         =   "△□□□的面积="
         End
      End
      Begin VB.Menu c_multi 
         Caption         =   "关于多边形的结论"
         Begin VB.Menu c_multi1 
            Caption         =   "□□□□是平行四边形"
         End
         Begin VB.Menu c_multi2 
            Caption         =   "□□□□是正方形"
         End
         Begin VB.Menu c_multi3 
            Caption         =   "□□□□是长方形"
         End
         Begin VB.Menu c_multi4 
            Caption         =   "□□□□是等腰梯形"
         End
         Begin VB.Menu c_multi5 
            Caption         =   "□□□□是菱形"
         End
         Begin VB.Menu c_multi6 
            Caption         =   "四边形□□□□的周长="
         End
         Begin VB.Menu c_multi7 
            Caption         =   "四边形□□□□的面积="
         End
      End
      Begin VB.Menu c_circle 
         Caption         =   "关于圆的结论"
         Begin VB.Menu c_circle1 
            Caption         =   "□、□、□、□四点共圆"
         End
         Begin VB.Menu c_circle2 
            Caption         =   "直线□□与⊙□_相切于□"
         End
      End
      Begin VB.Menu c_cal 
         Caption         =   "计算题"
         Begin VB.Menu c_cal2 
            Caption         =   "求∠□□□的大小"
         End
         Begin VB.Menu c_cal1 
            Caption         =   "求线段□□的长"
         End
         Begin VB.Menu c_cal3 
            Caption         =   "求线段□□，□□的比"
         End
         Begin VB.Menu c_cal4 
            Caption         =   "求△□□□的面积"
         End
         Begin VB.Menu c_cal5 
            Caption         =   "求四边形□□□□的面积"
         End
         Begin VB.Menu c_calD 
            Caption         =   "求△□□□与△□□□的面积比"
         End
         Begin VB.Menu c_cal6 
            Caption         =   "求⊙□_的面积"
         End
         Begin VB.Menu c_cal7 
            Caption         =   "求扇形□□□的面积"
         End
         Begin VB.Menu c_cal8 
            Caption         =   "求△□□□的周长"
         End
         Begin VB.Menu c_calA 
            Caption         =   "求四边形□□□□的周长"
         End
         Begin VB.Menu c_cal9 
            Caption         =   "求⊙□_的周长"
         End
         Begin VB.Menu c_calB 
            Caption         =   "求一般线段表达式的值"
         End
         Begin VB.Menu c_calC 
            Caption         =   "求以_和_为根的一元二次方程"
         End
         Begin VB.Menu c_calF 
            Caption         =   "求_的极□值"
         End
      End
      Begin VB.Menu c_choose 
         Caption         =   "选择题"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mea_and_cal 
      Caption         =   "测量[&M]"
      Begin VB.Menu set_for_measure 
         Caption         =   "设定测量标准"
         HelpContextID   =   57650
         Begin VB.Menu set_ruler 
            Caption         =   "设置标尺"
         End
         Begin VB.Menu set_length 
            Caption         =   "设线段长"
         End
         Begin VB.Menu set_dis_p_line 
            Caption         =   "设点到直线距离"
         End
         Begin VB.Menu set_area__of_polygon 
            Caption         =   "设多边形面积"
         End
      End
      Begin VB.Menu measure 
         Caption         =   "测量"
         Begin VB.Menu length 
            Caption         =   "测量长度"
            HelpContextID   =   57661
         End
         Begin VB.Menu distance_p_line 
            Caption         =   "测量点到直线的距离"
            HelpContextID   =   57662
         End
         Begin VB.Menu eara 
            Caption         =   "测量多边形面积"
            HelpContextID   =   57663
         End
         Begin VB.Menu angle 
            Caption         =   "测量角度"
            HelpContextID   =   57664
         End
         Begin VB.Menu image_of_function 
            Caption         =   "显示相关量的函数图像"
         End
      End
   End
   Begin VB.Menu M_run 
      Caption         =   "解题[&R]"
      Index           =   0
      Begin VB.Menu chose_law 
         Caption         =   "选择推理规则"
         HelpContextID   =   57670
      End
      Begin VB.Menu display_style 
         Caption         =   "选择显示方式"
         HelpContextID   =   57671
         Begin VB.Menu method0 
            Caption         =   "整体显示"
            Checked         =   -1  'True
         End
         Begin VB.Menu method1 
            Caption         =   "单步显示"
         End
      End
      Begin VB.Menu solve 
         Caption         =   "解题"
         HelpContextID   =   57673
      End
      Begin VB.Menu method 
         Caption         =   "自动解题"
         HelpContextID   =   57673
         Visible         =   0   'False
      End
      Begin VB.Menu method2 
         Caption         =   "交互解题"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu examp 
      Caption         =   "例题[&L]"
      HelpContextID   =   57674
   End
   Begin VB.Menu dbase 
      Caption         =   "信息库[&I]"
      HelpContextID   =   57680
      Begin VB.Menu angle_inform 
         Caption         =   "有关角的信息"
         Enabled         =   0   'False
         Begin VB.Menu two_angle 
            Caption         =   "两角和"
            Enabled         =   0   'False
         End
         Begin VB.Menu angle_relation 
            Caption         =   "两角比"
            Enabled         =   0   'False
         End
         Begin VB.Menu three_angle 
            Caption         =   "三角和"
            Enabled         =   0   'False
         End
         Begin VB.Menu sum_two_angle_right 
            Caption         =   "互余角"
            Enabled         =   0   'False
         End
         Begin VB.Menu sum_two_angle_pi 
            Caption         =   "互补角"
            Enabled         =   0   'False
         End
         Begin VB.Menu eangle 
            Caption         =   "相等角"
            Enabled         =   0   'False
         End
         Begin VB.Menu yizhiA 
            Caption         =   "已知角"
            Enabled         =   0   'False
         End
         Begin VB.Menu right_angle 
            Caption         =   "直　角"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu inform_line 
         Caption         =   "有关直线的信息"
         Enabled         =   0   'False
         Begin VB.Menu paral 
            Caption         =   "平行直线"
            Enabled         =   0   'False
         End
         Begin VB.Menu verti 
            Caption         =   "垂直直线"
            Enabled         =   0   'False
         End
         Begin VB.Menu three_point_on_line 
            Caption         =   "三点共线"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu inform_segment 
         Caption         =   "有关线段的信息"
         Enabled         =   0   'False
         Begin VB.Menu length_of_segment 
            Caption         =   "线段长"
            Enabled         =   0   'False
         End
         Begin VB.Menu two_line_value 
            Caption         =   "两线段和"
            Enabled         =   0   'False
         End
         Begin VB.Menu eline 
            Caption         =   "相等线段"
            Enabled         =   0   'False
         End
         Begin VB.Menu relation 
            Caption         =   "线段比"
            Enabled         =   0   'False
         End
         Begin VB.Menu re_line 
            Caption         =   "比例线段"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu inform_circle 
         Caption         =   "有关圆的信息"
         Enabled         =   0   'False
         Begin VB.Menu four_point_on_circle 
            Caption         =   "四点共圆"
            Enabled         =   0   'False
         End
         Begin VB.Menu area_of_circle 
            Caption         =   "圆的面积"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu inform_triangle 
         Caption         =   "有关三角形的信息"
         Enabled         =   0   'False
         Begin VB.Menu total_equal_triangle 
            Caption         =   "全等三角形"
            Enabled         =   0   'False
         End
         Begin VB.Menu similar_triangle 
            Caption         =   "相似三角形"
            Enabled         =   0   'False
         End
         Begin VB.Menu area_of_triangle 
            Caption         =   "三角形的面积"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu infrom_polygon 
         Caption         =   "有关多边形的信息"
         Enabled         =   0   'False
         Begin VB.Menu sp_four_sides 
            Caption         =   "特殊四边形"
            Enabled         =   0   'False
         End
         Begin VB.Menu area_of_polygon 
            Caption         =   "四边形面积"
            Enabled         =   0   'False
         End
         Begin VB.Menu epolygon 
            Caption         =   "正多边形"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu Help 
      Caption         =   "帮助[&H]"
      HelpContextID   =   57690
      Begin VB.Menu about 
         Caption         =   "关于《DDS-平面几何》"
      End
      Begin VB.Menu tool 
         Caption         =   "常用工具"
         Begin VB.Menu calc 
            Caption         =   "计算器"
         End
      End
      Begin VB.Menu knowl 
         Caption         =   "初中平面几何知识要点"
      End
      Begin VB.Menu index 
         Caption         =   "帮助主题[&I]"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public frm_sh As frmSplash
Option Explicit
Dim dir1$
Dim test_no%
Public time11_display_type As Byte
Dim time12_display_type As Byte
Private Sub about_Click()
frmAbout.Show
' Call MsgBox("    《欧几理德-平面几何》由中国科学院成都计算机应用研究所丁孙荭研究员和成都信息工程学院教师丁思捷研究开发。", _
    vbDefaultButton2, "", 0, 0)
End Sub
Public Sub Set_Mune_Item()
If regist_data.language = 0 Then
   regist_data.language = 1
End If
Mdiform1_caption = LoadResString_(110, "") & App.Major & "." & App.Minor & "." & App.Revision
'Mdiform1_caption = Me.Caption
Me.File.Caption = LoadResString_(100, "") '文件
Me.Open.Caption = LoadResString_(105, "") '打开文件
Me.about.Caption = LoadResString_(2275, "\\1\" + LoadResString_(110, "")) '关于
Me.save.Caption = LoadResString_(115, "") '存盘
Me.save_as.Caption = LoadResString_(120, "") '另存为
Me.savee.Caption = LoadResString_(125, "") '存例
Me.mprint.Caption = LoadResString_(150, "") '打印
Me.exit.Caption = LoadResString_(135, "") '退出
Me.edit.Caption = LoadResString_(140, "") '编辑
Me.set_mode.Caption = LoadResString_(145, "") '设置
Me.new.Caption = LoadResString_(160, "") '新建
Me.re_name.Caption = LoadResString_(165, "") '重命名
Me.re_name_all.Caption = LoadResString_(170, "") '重命名所有点
Me.re_name_one.Caption = LoadResString_(175, "") '重命名个别点
Me.remove_point_.Caption = LoadResString_(180, "") '删除点
Me.note.Caption = LoadResString_(185, "") '注释
Me.delete_last_op.Caption = LoadResString_(190, "") '撤消最后一步操作
Me.set_p.Caption = LoadResString_(195, "") '设置作图板属性
Me.draw.Caption = LoadResString_(200, "") '作图
Me.porine.Caption = LoadResString_(205, "") '作点和线
Me.pandline.Caption = Me.porine.Caption '
Me.midpoint.Caption = LoadResString_(210, "") '作线段中点
Me.ratio_point.Caption = LoadResString_(215, "") '作线段的定比分点
Me.line_given_length.Caption = LoadResString_(220, "") '作定长线段
Me.equal_line.Caption = LoadResString_(225, "") '作等长线段
Me.equal_angle_line.Caption = LoadResString_(230, "") '作角平分线
Me.draw_circle.Caption = LoadResString_(235, "") '画圆
Me.d_circle.Caption = LoadResString_(235, "") '画圆
'Me.d_circle_without_center.Caption = LoadResString_(240, "") '画无心圆
Me.tangent_line_point_to_circle.Caption = LoadResString_(245, "") '画切线
Me.tangent_of_two_circles.Caption = LoadResString_(4305, "") '画两圆的公切线
'Me.tangent_circle.Caption = LoadResString_(250, "") '画相切圆
Me.paralandverti0.Caption = LoadResString_(255, "") '画平行（垂直）线
Me.paralandverti.Caption = LoadResString_(255, "") '画平行（垂直）线
Me.verti_mid_line.Caption = LoadResString_(265, "") '画垂直平分线
Me.E_polygon.Caption = LoadResString_(270, "") '作正多边形
Me.e_polygon3.Caption = LoadResString_(275, "") '正三角形
Me.e_polygon4.Caption = LoadResString_(280, "") '正四边形
Me.e_polygon5.Caption = LoadResString_(285, "") '正五边形
Me.e_polygon6.Caption = LoadResString_(290, "") '正六边形
Me.move_picture.Caption = LoadResString_(295, "") '作动态图
Me.set_view_point.Caption = LoadResString_(300, "") '设置观察点
Me.set_opera_type.Caption = LoadResString_(305, "") '设置操作方式
Me.by_hand.Caption = LoadResString_(310, "") '手动
Me.auto.Caption = LoadResString_(315, "") '自动
Me.change_picture.Caption = LoadResString_(320, "") '图形变换
Me.set_picture_for_change.Caption = LoadResString_(325, "") '设置变换图形
Me.set_change_line.Caption = LoadResString_(330, "") '设置变换直线
Me.set_polygon_for_change.Caption = LoadResString_(335, "") '设置变换多边形
Me.set_circle_for_change.Caption = LoadResString_(340, "") '设置变换圆
Me.set_change_type.Caption = LoadResString_(345, "") '设置变换方式
Me.move_part.Caption = LoadResString_(350, "") '平移
Me.turn_part.Caption = LoadResString_(355, "") '旋转
Me.turn_over.Caption = LoadResString_(360, "") '翻转
Me.fd_sx_p.Caption = LoadResString_(365, "") '放大缩小
Me.zhouduicheng.Caption = LoadResString_(370, "") '轴对称
Me.zhongxinduichen.Caption = LoadResString_(375, "") '中心对称
Me.initial_picture.Caption = LoadResString_(380, "") '恢复原图
Me.delete_picture_change.Caption = LoadResString_(385, "") '撤消图形变换
Me.judge.Caption = LoadResString_(390, "") '判断
Me.chose1.Caption = LoadResString_(395, "") '选择一
Me.chose2.Caption = LoadResString_(400, "") '选择二
Me.chose3.Caption = LoadResString_(405, "") '选择三
Me.Inputcond.Caption = LoadResString_(410, "") '条件
Me.s1.Caption = LoadResString_(415, "") '一般条件
Me.S12.Caption = set_display_two_icon & "=" & set_display_two_icon 'XX=XX
Me.S4A.Caption = simple_display_string(LoadResString_from_inpcond(-24, "\\0\\" & global_icon_char & _
                                                 "\\1\\" & global_icon_char & _
                                                 "\\2\\" & global_icon_char & _
                                                 "\\3\\" & global_icon_char)) '弧XXX=弧XXX
Me.S14.Caption = set_display_angle0(set_display_three_icon) & "=" & set_display_angle0(set_display_three_icon) '〈XXX=〈XXX
Me.S15.Caption = set_display_angle0(set_display_three_icon) & "=_" ' 〈XXX=_
Me.s16.Caption = set_display_two_icon & "=_" 'XX=_
Me.S17.Caption = set_display_two_icon & ":" & set_display_two_icon & "=_" 'XX:XX=_
Me.s18.Caption = set_display_two_icon & "+" & set_display_two_icon & "=_" 'XX+XX=_
Me.s19.Caption = set_display_two_icon & ":" & set_display_two_icon & "=" & _
                 set_display_two_icon & ":" & set_display_two_icon 'XX:XX=XX：XX
Me.s1A.Caption = set_display_two_icon & "=" & set_display_two_icon & ":" & set_display_two_icon 'XX=XX+XX
Me.s1B.Caption = set_display_angle0(set_display_three_icon) & ":" & set_display_angle0(set_display_three_icon) & "=_" '〈XXX:<***=_
Me.s1D.Caption = simple_display_string( _
                        LoadResString_from_inpcond(-50, "\\0\\" & global_icon_char & _
                                                  "\\1\\" & global_icon_char & _
                                                  "\\2\\" & global_icon_char & _
                                                  "\\3\\" & global_icon_char)) '**是<***的平分线
Me.s1C.Caption = LoadResString_(480, "") '线段的一般关系式
Me.s1G.Caption = simple_display_string(LoadResString_from_inpcond(22, "\\0\\_\\1\\_")) '*,*是一元二次方程_的两个根
Me.S2.Caption = LoadResString_(490, "") '直线和直线上的交点
Me.S21.Caption = simple_display_string( _
                           LoadResString_from_inpcond(1, "\\0\\" & global_icon_char & "\\1\\" & _
                                                       global_icon_char & "\\2\\" & global_icon_char)) '直线**上任取一点
Me.S22.Caption = simple_display_string( _
                           LoadResString_from_inpcond(2, "\\0\\" & global_icon_char & _
                                               "\\1\\" & global_icon_char & _
                                               "\\2\\" & global_icon_char & _
                                               "\\3\\" & global_icon_char)) '**平行**
Me.S23.Caption = set_display_two_icon & LoadResString_(1405, "") & set_display_two_icon 'xx垂直xx
Me.S24.Caption = simple_display_string( _
                           LoadResString_from_inpcond(4, "\\0\\" & global_icon_char & _
                                               "\\1\\" & global_icon_char & _
                                               "\\2\\" & global_icon_char)) '在**的垂直平分线上任取一点
Me.S31.Caption = simple_display_string( _
                           LoadResString_from_inpcond(5, "\\0\\" & global_icon_char & _
                                               "\\1\\" & global_icon_char & _
                                               "\\2\\" & global_icon_char)) '取线段的中点
Me.S32.Caption = simple_display_string(LoadResString_from_inpcond(6, "\\0\\" & global_icon_char & _
                                       "\\1\\" & global_icon_char & _
                                       "\\2\\" & global_icon_char & _
                                       "\\3\\" & "_")) '*是线段**上分比为_的分点
Me.S34.Caption = simple_display_string(LoadResString_from_inpcond(14, "\\0\\" & global_icon_char & _
                                                "\\1\\" & global_icon_char & _
                                                "\\2\\" & global_icon_char & _
                                                "\\3\\" & global_icon_char)) '过*作线段**的垂直平分线垂足为*
Me.S4.Caption = LoadResString_(530, "") '圆和关于圆的点线
Me.S41.Caption = simple_display_string( _
                              LoadResString_from_inpcond(7, "\\0\\" & global_icon_char & _
                                               "\\1\\" & "_" & _
                                               "\\2\\" & global_icon_char)) '圆*_上任取一点
Me.S42.Caption = simple_display_string( _
                          LoadResString_from_inpcond(8, "\\0\\" & global_icon_char & _
                                               "\\1\\" & global_icon_char & _
                                               "\\2\\" & global_icon_char & _
                                               "\\3\\" & global_icon_char & _
                                               "\\4\\" & "_")) '过***作圆
Me.S4C.Caption = simple_display_string(LoadResString_from_inpcond(12, _
                  "\\0\\" & global_icon_char & "\\1\\_" & _
                  "\\2\\" & global_icon_char & "\\3\\_" & _
                  "\\4\\" & global_icon_char)) '圆和圆相切与*
Me.S45.Caption = simple_display_string(LoadResString_from_inpcond(-33, _
                       "\\0\\" & global_icon_char & _
                       "\\1\\" & global_icon_char & _
                       "\\2\\_" & "\\3\\" & global_icon_char & _
                       "\\0\\" & global_icon_char)) '过*作圆的切线**
Me.S5.Caption = LoadResString_(555, "") '三角形的心
Me.S51.Caption = simple_display_string(LoadResString_from_inpcond(18, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char)) '*是三角形的重心
Me.S52.Caption = simple_display_string(LoadResString_from_inpcond(19, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char)) '*是三角形外接圆的圆心
Me.S53.Caption = simple_display_string(LoadResString_from_inpcond(20, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char)) '*是三角形的垂心
Me.S54.Caption = simple_display_string(LoadResString_from_inpcond(21, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char)) '*是三角形内切圆的圆心
Me.S6.Caption = LoadResString_(580, "") '多边形
Me.S61.Caption = simple_display_string(LoadResString_from_inpcond(-20, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char)) '任意三角形***
Me.S6E.Caption = simple_display_string(LoadResString_from_inpcond(62, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & "\\3\\_")) '三角形***的周长
Me.S6F.Caption = simple_display_string(LoadResString_from_inpcond(63, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & "\\3\\_")) '三角形***的面积
Me.S6I.Caption = simple_display_string(LoadResString_from_inpcond(-37, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char & _
                          "\\4\\" & global_icon_char & _
                          "\\5\\" & global_icon_char)) '三角形***全等于三角形***
Me.S6J.Caption = LoadResString_from_inpcond(-36, set_display_triangle0(set_display_three_icon, 1, 0) & _
                                 set_display_triangle0(set_display_three_icon, 1, 3)) '三角形***相似于三角形***
Me.S62.Caption = simple_display_string(LoadResString_from_inpcond(-19, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char)) '任意四边形****
Me.S6G.Caption = simple_display_string(LoadResString_from_inpcond(64, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char & "\\4\\_")) '四边形****的周长=_
Me.S6H.Caption = simple_display_string(LoadResString_from_inpcond(65, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char & "\\4\\_")) '四边形****的面积=_
Me.S63.Caption = simple_display_string(LoadResString_from_inpcond(-18, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char)) '***是等腰三角形
Me.S64.Caption = simple_display_string(LoadResString_from_inpcond(-17, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char)) '***是等腰直角三角形
Me.S65.Caption = simple_display_string(LoadResString_from_inpcond(-16, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char)) '***是等边三角形
Me.S66.Caption = LoadResString_from_inpcond(48, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char) '****是梯形
Me.S67.Caption = simple_display_string(LoadResString_from_inpcond(-14, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char)) '****是等腰梯形
Me.S68.Caption = simple_display_string(LoadResString_from_inpcond(-13, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char)) '****是长方形
Me.S69.Caption = simple_display_string(LoadResString_from_inpcond(-12, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char)) '****是正方形
Me.S6A.Caption = LoadResString_from_inpcond(-11, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char) '****是平行四边形
Me.S6B.Caption = simple_display_string(LoadResString_from_inpcond(-10, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char)) '****是菱形
Me.S6C.Caption = LoadResString_from_inpcond(-9, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char & _
                          "\\4\\" & global_icon_char) '*****是正五边形
Me.S6D.Caption = LoadResString_from_inpcond(-8, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char & _
                          "\\4\\" & global_icon_char & _
                          "\\5\\" & global_icon_char) '******是正六边形
Me.conclusion.Caption = LoadResString_(680, "") '结论
Me.C_line.Caption = LoadResString_(685, "") '关于直线的结论
Me.c_line1.Caption = simple_display_string(LoadResString_from_inpcond(24, "\\0\\" & global_icon_char & _
                                         "\\1\\" & global_icon_char & _
                                         "\\2\\" & global_icon_char)) '***三点共线
Me.c_line2.Caption = simple_display_string( _
                    LoadResString_from_inpcond(39, "\\0\\" & global_icon_char & _
                                      "\\1\\" & global_icon_char & _
                                      "\\2\\" & global_icon_char & _
                                      "\\3\\" & global_icon_char & _
                                      "\\4\\" & global_icon_char & _
                                      "\\5\\" & global_icon_char)) '**，**，**三线共点
Me.c_line3.Caption = simple_display_string( _
                       LoadResString_from_inpcond(2, "\\0\\" & global_icon_char & _
                                       "\\1\\" & global_icon_char & _
                                       "\\2\\" & global_icon_char & _
                                       "\\3\\" & global_icon_char)) '**平行**重复s22
Me.c_line4.Caption = simple_display_string( _
                       LoadResString_from_inpcond(3, "\\0\\" & global_icon_char & _
                                       "\\1\\" & global_icon_char & _
                                       "\\2\\" & global_icon_char & _
                                       "\\3\\" & global_icon_char)) '**垂直**s23
Me.c_line5.Caption = simple_display_string( _
                       LoadResString_from_inpcond(29, "\\0\\" & global_icon_char & _
                                         "\\1\\" & global_icon_char & _
                                          "\\2\\" & global_icon_char)) '点*位于**的垂直平分线上
Me.cl1.Caption = LoadResString_(705, "") '关于线段的结论
Me.c_linel1_0.Caption = set_display_two_icon & "=_" '**=-
Me.c_linel1_1.Caption = simple_display_string( _
                         LoadResString_from_inpcond(26, "\\0\\" & global_icon_char & _
                                         "\\1\\" & global_icon_char & _
                                         "\\2\\" & global_icon_char)) '*是线段**的中点
Me.c_line1_2.Caption = set_display_two_icon & ":" & set_display_two_icon & "=_" '**：**=-
Me.c_line1_3.Caption = set_display_two_icon & "=" & set_display_two_icon '
Me.c_line1_4.Caption = set_display_two_icon & ":" & set_display_two_icon & "=" & _
                       set_display_two_icon & ":" & set_display_two_icon '
Me.c_line1_5.Caption = LoadResString_(715, "") '
Me.c_line1_6.Caption = LoadResString_(720, "\\1\\_") '
Me.c_angle.Caption = LoadResString_(725, "") '
Me.c_angle0.Caption = set_display_angle0(set_display_three_icon) & "=_" '
Me.c_angle1.Caption = set_display_angle0(set_display_three_icon) & "=" & set_display_angle0(set_display_three_icon) '
Me.c_angle2.Caption = set_display_angle0(set_display_three_icon) & ":" & set_display_angle0(set_display_three_icon) & "=_" '
Me.c_angle3.Caption = set_display_angle0(set_display_three_icon) & "+" & set_display_angle0(set_display_three_icon) & "=_" '
Me.c_angle4.Caption = set_display_angle0(set_display_three_icon) & "=" & set_display_angle0(set_display_three_icon) & "+" & _
                     set_display_angle0(set_display_three_icon) '
Me.c_triangle.Caption = LoadResString_(740, "") '
Me.c_triangle1.Caption = simple_display_string(LoadResString_from_inpcond(-37, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char & _
                          "\\4\\" & global_icon_char & _
                          "\\5\\" & global_icon_char)) '
Me.c_triangle2.Caption = simple_display_string(LoadResString_from_inpcond(-36, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char & _
                          "\\4\\" & global_icon_char & _
                          "\\5\\" & global_icon_char)) '
Me.c_triangle8.Caption = simple_display_string(LoadResString_from_inpcond(68, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char & _
                          "\\4\\" & global_icon_char & _
                          "\\5\\" & global_icon_char)) '
Me.c_triangle3.Caption = simple_display_string(LoadResString_from_inpcond(-16, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char)) '
Me.c_triangle5.Caption = simple_display_string(LoadResString_from_inpcond(-18, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char)) '
Me.c_triangle4.Caption = simple_display_string(LoadResString_from_inpcond(-17, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char)) '
Me.c_triangle6.Caption = simple_display_string(LoadResString_from_inpcond(62, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & "\\3\\_")) '
Me.c_triangle7.Caption = simple_display_string(LoadResString_from_inpcond(63, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & "\\3\\_")) '
Me.c_multi.Caption = LoadResString_(750, "") '
Me.c_multi1.Caption = LoadResString_from_inpcond(-11, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char) '
Me.c_multi2.Caption = simple_display_string(LoadResString_from_inpcond(-12, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char)) '
Me.c_multi3.Caption = simple_display_string(LoadResString_from_inpcond(-13, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char)) '
Me.c_multi4.Caption = simple_display_string(LoadResString_from_inpcond(-14, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char)) '
Me.c_multi5.Caption = simple_display_string(LoadResString_from_inpcond(-11, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char)) '
Me.c_multi6.Caption = simple_display_string(LoadResString_from_inpcond(64, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char & "\\4\\_")) '
Me.c_multi7.Caption = simple_display_string(LoadResString_from_inpcond(65, _
                          "\\0\\" & global_icon_char & _
                          "\\1\\" & global_icon_char & _
                          "\\2\\" & global_icon_char & _
                          "\\3\\" & global_icon_char & "\\4\\_")) '
Me.c_circle.Caption = LoadResString_(755, "") '
Me.c_circle1.Caption = LoadResString_from_inpcond(23, "\\0\\" & global_icon_char & _
                                           "\\1\\" & global_icon_char & _
                                           "\\2\\" & global_icon_char & _
                                           "\\3\\" & global_icon_char) '
Me.c_circle2.Caption = simple_display_string(LoadResString_from_inpcond(47, _
                                           "\\0\\" & global_icon_char & _
                                           "\\1\\" & global_icon_char & _
                                           "\\2\\" & global_icon_char & "\\3\\_" & _
                                           "\\4\\" & global_icon_char)) '
Me.c_cal.Caption = LoadResString_(770, "") '
Me.c_cal2.Caption = LoadResString_from_inpcond(36, _
                                           "\\0\\" & global_icon_char & _
                                           "\\1\\" & global_icon_char & _
                                           "\\2\\" & global_icon_char) '
Me.c_cal1.Caption = LoadResString_from_inpcond(35, _
                                        "\\0\\" & global_icon_char & _
                                        "\\1\\" & global_icon_char) '
Me.c_cal3.Caption = LoadResString_from_inpcond(37, _
                                           "\\0\\" & global_icon_char & _
                                           "\\1\\" & global_icon_char & _
                                           "\\2\\" & global_icon_char & _
                                           "\\3\\" & global_icon_char) '
Me.c_cal4.Caption = LoadResString_from_inpcond(56, _
                                           "\\0\\" & global_icon_char & _
                                           "\\1\\" & global_icon_char & _
                                           "\\2\\" & global_icon_char) '
Me.c_cal5.Caption = LoadResString_from_inpcond(65, _
                                           "\\0\\" & global_icon_char & _
                                           "\\1\\" & global_icon_char & _
                                           "\\2\\" & global_icon_char & _
                                           "\\3\\" & global_icon_char) '
Me.c_calD.Caption = simple_display_string( _
                    LoadResString_from_inpcond(69, set_display_triangle0(set_display_three_icon, 1, 0) & _
                                                   set_display_triangle0(set_display_three_icon, 1, 3))) '
Me.c_cal6.Caption = simple_display_string(LoadResString_from_inpcond(58, _
                                           "\\0\\" & global_icon_char & _
                                           "\\1\\_")) '
Me.c_cal7.Caption = LoadResString_from_inpcond(59, _
                                           "\\0\\" & global_icon_char & _
                                           "\\1\\" & global_icon_char & _
                                           "\\2\\" & global_icon_char) '
Me.c_cal8.Caption = simple_display_string(LoadResString_from_inpcond(60, _
                      set_display_triangle0(set_display_three_icon, 1, 0))) '
Me.c_calA.Caption = LoadResString_from_inpcond(64, _
                                           "\\0\\" & global_icon_char & _
                                           "\\1\\" & global_icon_char & _
                                           "\\2\\" & global_icon_char & _
                                           "\\3\\" & global_icon_char) '
Me.c_cal9.Caption = simple_display_string(LoadResString_from_inpcond(58, "\\0\\" + global_icon_char & "\\1\\_")) '
Me.c_calB.Caption = simple_display_string(LoadResString_from_inpcond(50, "\\0\\_")) '
Me.c_calC.Caption = simple_display_string(LoadResString_from_inpcond(67, "\\0\\_\\1\\_")) '
Me.c_calF.Caption = simple_display_string(LoadResString_from_inpcond(74, "\\0\\_")) '
Me.c_choose.Caption = LoadResString_(850, "") & "[&M]" '
Me.mea_and_cal.Caption = LoadResString_(850, "") & "[&M]" '
Me.length.Caption = LoadResString_(855, "") '
Me.set_for_measure.Caption = LoadResString_(860, "") '
Me.set_ruler.Caption = LoadResString_(865, "") '
Me.set_length.Caption = LoadResString_(870, "") '
Me.set_dis_p_line.Caption = LoadResString_(875, "") '
Me.set_area__of_polygon.Caption = LoadResString_(880, "") '
Me.measure.Caption = LoadResString_(850, "") '
Me.distance_p_line.Caption = LoadResString_(890, "") '
Me.eara.Caption = LoadResString_(895, "") '
Me.angle.Caption = LoadResString_(900, "") '
Me.image_of_function.Caption = LoadResString_(905, "") '
Me.M_run.item(0).Caption = LoadResString_(910, "") + "[&R]" '
Me.chose_law.Caption = LoadResString_(915, "") '
Me.display_style.Caption = LoadResString_(920, "") '
Me.method0.Caption = LoadResString_(925, "")
Me.method1.Caption = LoadResString_(930, "") '
Me.solve.Caption = LoadResString_(910, "") '
Me.method.Caption = LoadResString_(425, "") '
Me.method2.Caption = LoadResString_(945, "") '
Me.examp.Caption = LoadResString_(950, "") '
Me.dbase.Caption = LoadResString_(955, "") & "[&I]" '
Me.angle_inform.Caption = LoadResString_(960, "") '
Me.two_angle.Caption = LoadResString_(965, "") '
Me.angle_relation.Caption = LoadResString_(970, "") '
Me.three_angle.Caption = LoadResString_(975, "") '
Me.sum_two_angle_right.Caption = LoadResString_(980, "") '
Me.sum_two_angle_pi.Caption = LoadResString_(985, "") '
Me.eangle.Caption = LoadResString_(990, "") '
Me.yizhiA.Caption = LoadResString_(995, "") '
Me.right_angle.Caption = LoadResString_(1000, "") '
Me.inform_line.Caption = LoadResString_(1005, "") '
Me.paral.Caption = LoadResString_(1010, "") '
Me.verti.Caption = LoadResString_(1015, "") '
Me.three_point_on_line.Caption = LoadResString_(1020, "") '
Me.inform_segment.Caption = LoadResString_(1025, "") '
Me.length_of_segment.Caption = LoadResString_(1030, "") '
Me.two_line_value.Caption = LoadResString_(1035, "") '
Me.eline.Caption = LoadResString_(1040, "") '
Me.relation.Caption = LoadResString_(1045, "") '
Me.re_line.Caption = LoadResString_(1050, "") '
Me.inform_circle.Caption = LoadResString_(1055, "") '
Me.four_point_on_circle.Caption = LoadResString_(1060, "") '
Me.area_of_circle.Caption = LoadResString_(1065, "") '
Me.inform_triangle.Caption = LoadResString_(1070, "") '
Me.total_equal_triangle.Caption = LoadResString_(1075, "") '
Me.similar_triangle.Caption = LoadResString_(1080, "") '
Me.area_of_triangle.Caption = LoadResString_(1085, "") '
Me.infrom_polygon.Caption = LoadResString_(1090, "") '
Me.sp_four_sides.Caption = LoadResString_(1095, "") '
Me.area_of_polygon.Caption = LoadResString_(1100, "") '
Me.epolygon.Caption = LoadResString_(1105, "") '
Me.Help.Caption = LoadResString_(1105, "") '
Me.tool.Caption = LoadResString_(1115, "") '
Me.calc.Caption = LoadResString_(1120, "") '
Me.knowl.Caption = LoadResString_(1125, "") '
Me.index.Caption = LoadResString_(1130, "") '
Me.Toolbar1.Buttons(2).ToolTipText = LoadResString_(1925, "") '新问题
Me.Toolbar1.Buttons(3).ToolTipText = LoadResString_(105, "") '打开文件
Me.Toolbar1.Buttons(4).ToolTipText = LoadResString_(115, "") '存盘
Me.Toolbar1.Buttons(5).ToolTipText = LoadResString_(150, "") '打印
Me.Toolbar1.Buttons(7).ToolTipText = LoadResString_(205, "") '点和线
Me.Toolbar1.Buttons(8).ToolTipText = LoadResString_(235, "") '画圆
Me.Toolbar1.Buttons(9).ToolTipText = LoadResString_(255, "") '平行垂直
Me.Toolbar1.Buttons(10).ToolTipText = LoadResString_(270, "") '画多边形
Me.Toolbar1.Buttons(11).ToolTipText = LoadResString_(295, "") '画动态图
Me.Toolbar1.Buttons(12).ToolTipText = LoadResString_(320, "") '图形变换
Me.Toolbar1.Buttons(13).ToolTipText = LoadResString_(850, "") '测量
Me.Toolbar1.Buttons(14).ToolTipText = LoadResString_(860, "") '设置标尺
Me.Toolbar1.Buttons(15).ToolTipText = LoadResString_(885, "") '提问
Me.Toolbar1.Buttons(16).ToolTipText = LoadResString_(935, "") '修改输入
Me.Toolbar1.Buttons(17).ToolTipText = LoadResString_(4095, "") '暂停控制
Me.Toolbar1.Buttons(18).ToolTipText = LoadResString_(4295, "") '终止解题
Me.Toolbar1.Buttons(19).ToolTipText = LoadResString_(1935, "") '解题
Me.Toolbar1.Buttons(20).ToolTipText = LoadResString_(1930, "") '标示笔
Me.Toolbar1.Buttons(21).ToolTipText = LoadResString_(4300, "") '撤消最后一步操作
End Sub

Private Sub angle_Click()
If run_type = 0 Then
operator = "measure"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 4
old_operator = operator
Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2245, "")
'Draw_form.SetFocus
'protect_munu = 1
End If
End Sub

Private Sub angle_relation_Click() '角的关系
Call set_inform_list(LoadResString_(955, "") & ":" & angle_relation.Caption)
inform.Picture1.Cls
Call set_information_list(angle_relation_)
End Sub
Private Sub area_of_circle_Click() '
Call set_inform_list(LoadResString_(955, "") & ":" & area_of_circle.Caption)
inform.Picture1.Cls
Call set_information_list(area_of_circle_)
End Sub
Private Sub area_of_polygon_Click() '
Call set_inform_list(LoadResString_(955, "") & ":" & area_of_polygon.Caption)
inform.Picture1.Cls
Call set_information_list(area_of_polygon_)
End Sub
Private Sub area_of_triangle_Click() '
Call set_inform_list(LoadResString_(955, "") & ":" & area_of_triangle)
inform.Picture1.Cls
Call set_information_list(area_of_triangle_)
End Sub
Private Sub auto_Click() '
If run_type > 0 Then  ' 已作图形变换
 Exit Sub
End If
'If operator = "move_point" And _
'     last_conditions.last_cond(1).last_view_point_no > 0 Then
'      operator = "move_point_"
'ElseIf last_conditions.last_cond(1).last_view_point_no = 0 Then
'  Draw_form.Picture1.visible = False
  operator = "move_point"
   'Set C_curve = New curve_Class
   chose1.Caption = LoadResString_(2290, "") '设置直线轨道
   chose2.Caption = LoadResString_(2295, "") '设置圆形轨道
   chose3.Caption = LoadResString_(2300, "") '"自定义运动轨道"
   chose3.visible = True
   Call remove_uncomplete_operat(old_operator)
   list_type_for_draw = 2
   old_operator = operator
     MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2305, "")
         '  Draw_form.SetFocus
'End If
End Sub
Private Sub by_hand_Click()
If change_pic Then '正作图形变换
Exit Sub
End If
'If operator = "move_point" And _
'     last_conditions.last_cond(1).last_view_point_no > 0 Then
'      operator = "move_point_"
'ElseIf last_conditions.last_cond(1).last_view_point_no = 0 Then
'  Draw_form.Picture1.visible = False
  operator = "move_point"
  Call remove_uncomplete_operat(old_operator)
  list_type_for_draw = 1
  old_operator = operator
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2220, "")
  'Draw_form.SetFocus
 ' protect_munu = 1
'End If
End Sub

Private Sub c_angle0_Click()
wenti_type = 0
Call input_sentence_y(2, 53, 1)
End Sub
Private Sub c_angle1_Click()
wenti_type = 0
Call input_sentence_y(2, 30, 1)
End Sub
Private Sub c_angle2_Click()
wenti_type = 0
Call input_sentence_y(2, 55, 1)
End Sub
Private Sub c_angle3_Click()
wenti_type = 0
Call input_sentence_y(2, 52, 1)
End Sub
Private Sub c_angle4_Click()
wenti_type = 0
Call input_sentence_y(2, 51, 1)
End Sub

Private Sub c_cal1_Click()
Call input_sentence_y(2, 35, 1)
If last_conclusion = 0 Then
 wenti_type = 1
End If
End Sub
Private Sub c_cal2_Click()
Call input_sentence_y(2, 36, 1)
If last_conclusion = 0 Then
 wenti_type = 1
End If
End Sub
Private Sub c_cal3_Click()
Call input_sentence_y(2, 37, 1)
If last_conclusion = 0 Then
 wenti_type = 1
End If
End Sub
Private Sub c_cal4_Click()
Call input_sentence_y(2, 56, 1)
If last_conclusion = 0 Then
 wenti_type = 1
End If
End Sub
Private Sub c_cal5_Click()
Call input_sentence_y(2, 57, 1)
If last_conclusion = 0 Then
 wenti_type = 1
End If
End Sub
Private Sub c_cal6_Click()
inp = 58
Call input_sentence_y(2, 58, 1)
If last_conclusion = 0 Then
 wenti_type = 1
End If
End Sub
Private Sub c_cal7_Click()
Call input_sentence_y(2, 59, 1)
If last_conclusion = 0 Then
 wenti_type = 1
End If
End Sub
Private Sub c_cal8_Click()
Call input_sentence_y(2, 60, 1)
If last_conclusion = 0 Then
 wenti_type = 1
End If
End Sub
Private Sub c_cal9_Click()
Call input_sentence_y(2, 61, 1)
If last_conclusion = 0 Then
 wenti_type = 1
End If
End Sub
Private Sub c_calA_Click()
Call input_sentence_y(2, 66, 1)
End Sub
Private Sub c_calB_Click()
wenti_type = 0
Call input_sentence_y(2, 50, 1)
End Sub

Private Sub c_calC_Click()
wenti_type = 0
Call input_sentence_y(2, 67, 1)
End Sub

Private Sub c_calD_Click()
Call input_sentence_y(2, 69, 1)
End Sub

Private Sub c_calF_Click()
wenti_type = 0
Call input_sentence_y(2, 74, 1)
End Sub

Private Sub c_choose_Click()
If event_statue = wait_for_draw_point Or _
     event_statue = ready Then
wenti_type = 1 '选择题
Wenti_form.Picture1.CurrentY = 20 * (C_display_wenti.m_last_input_wenti_no + 1)
Wenti_form.Picture1.CurrentX = 0
Wenti_form.Picture1.Print LoadResString_(435, "") + "(  )"
End If
End Sub

Private Sub c_circle1_Click()
wenti_type = 0
Call input_sentence_y(2, 23, 1)
End Sub

Private Sub c_circle2_Click()
wenti_type = 0
Call input_sentence_y(2, 47, 1)
End Sub


Private Sub c_line1_2_Click()
wenti_type = 0
Call input_sentence_y(2, 31, 1)
End Sub

Private Sub c_line1_3_Click()
wenti_type = 0
Call input_sentence_y(2, 25, 1)
End Sub

Private Sub c_line1_4_Click()
wenti_type = 0
Call input_sentence_y(2, 32, 1)
End Sub

Private Sub c_line1_5_Click()
wenti_type = 0
Call input_sentence_y(2, 38, 1)
End Sub

Private Sub c_line1_6_Click()
wenti_type = 0
Call input_sentence_y(2, 73, 1)
End Sub

Private Sub c_line1_Click()
wenti_type = 0
Call input_sentence_y(2, 24, 1)
End Sub

Private Sub c_line2_Click()
wenti_type = 0
Call input_sentence_y(2, 39, 1)
End Sub

Private Sub c_line3_Click()
wenti_type = 0
Call input_sentence_y(2, 27, 1)
End Sub

Private Sub c_line4_Click()
wenti_type = 0
Call input_sentence_y(2, 28, 1)
End Sub

Private Sub c_line5_Click()
wenti_type = 0
Call input_sentence_y(2, 29, 1)
End Sub

Private Sub c_linel1_0_Click()
wenti_type = 0
Call input_sentence_y(2, 54, 1)
End Sub

Private Sub c_linel1_1_Click()
wenti_type = 0
Call input_sentence_y(2, 26, 1)
End Sub

Private Sub c_multi1_Click()
wenti_type = 0
Call input_sentence_y(2, 45, 1)
End Sub

Private Sub c_multi2_Click()
wenti_type = 0
Call input_sentence_y(2, 44, 1)
End Sub

Private Sub c_multi3_Click()
wenti_type = 0
Call input_sentence_y(2, 43, 1)
End Sub

Private Sub c_multi4_Click()
wenti_type = 0
Call input_sentence_y(2, 49, 1)
End Sub

Private Sub c_multi5_Click()
wenti_type = 0
Call input_sentence_y(2, 46, 1)
End Sub

Private Sub c_multi6_Click()
Call input_sentence_y(2, 64, 1)
End Sub
Private Sub c_multi7_Click()
Call input_sentence_y(2, 65, 1)
End Sub
Private Sub c_tiangle1_Click()
wenti_type = 0
Call input_sentence_y(2, 34, 1)
End Sub

Private Sub c_tiangle2_Click()
wenti_type = 0
Call input_sentence_y(2, 33, 1)
End Sub

Private Sub c_triangle3_Click()
wenti_type = 0
Call input_sentence_y(2, 40, 1)
End Sub

Private Sub c_triangle4_Click()
wenti_type = 0
Call input_sentence_y(2, 42, 1)
End Sub
Private Sub c_triangle5_Click()
wenti_type = 0
Call input_sentence_y(2, 41, 1)
End Sub

Private Sub c_triangle6_Click()
Call input_sentence_y(2, 62, 1)
End Sub


Private Sub c_triangle7_Click()
Call input_sentence_y(2, 63, 1)
End Sub

Private Sub c_triangle8_Click()
Call input_sentence_y(2, 68, 1)
End Sub

Private Sub calc_Click()
Shell ("calc.exe"), 1
End Sub

Private Sub chose_law_Click()
If run_type > 0 And run_type < 11 Then
 Exit Sub
End If
If run_statue = 6 Then '12.10
'MDIForm1.add_point.Enabled = True
MDIForm1.method.Enabled = True
MDIForm1.method.Checked = True
MDIForm1.method2.Enabled = True
'MDIForm1.method3.Enabled = True
End If
Load ch_ruler
ch_ruler.Show
'End If
End Sub

Private Sub chose1_Click()
Dim ty As Boolean
If operator = "move_point" And list_type_for_draw = 2 Then
 list_type_for_draw = 3
ElseIf operator = "re_name" And list_type_for_draw = 3 Then
  If temp_point(0).no = last_conditions.last_cond(1).point_no Then
   If MsgBox(LoadResString_(1805, "\\1\\" + m_poi(temp_point(0).no).data(0).data0.name), vbYesNo, "", "", 0) = 6 Then
      Call remove_point(temp_point(0).no, display, 0)
   Else
  End If
  Else
  Call MsgBox(LoadResString_(1790, "\\1\\" + m_poi(temp_point(0).no).data(0).data0.name + _
                              "\\2\\" + m_poi(last_conditions.last_cond(1).point_no).data(0).data0.name), 0, "", "", 0)
  End If
End If
End Sub

Private Sub chose2_Click()
If operator = "move_point" And list_type_for_draw = 2 Then
 list_type_for_draw = 4
ElseIf operator = "re_name" And list_type_for_draw = 3 Then
operat_is_acting = False
End If
End Sub

Private Sub chose3_Click()
If operator = "move_point" And list_type_for_draw = 2 Then
 list_type_for_draw = 5
End If
End Sub

Private Sub conclusion_Click()
If wenti_type > 2 Then
 wenti_type = wenti_type - 2
End If
If event_statue <> input_prove_by_hand Then
'Call init_operat
event_statue = wait_for_input_condition
End If
End Sub

Private Sub d_circle_Click()
operator = "draw_circle"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 1
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2310, "")
'Draw_form.SetFocus
Me.ZOrder
End Sub

Private Sub d_circle_without_center_Click()
circle_with_center = False
operator = "draw_circle"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 2
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2310, "")
'Draw_form.SetFocus
'protect_munu = 1
End Sub

Private Sub dbase_Click()
'If run_type = 0 Or run_type > 10 Then
'Call clear_wenti_display
 event_statue = get_inform
  Call set_inform
'End If
End Sub

Private Sub delete_last_op_Click()
Call un_do
End Sub

Private Sub delete_picture_change_Click()
 Draw_form.Cls
' Call get_old_picture
 Call draw_again1(Draw_form)
 last_conditions.last_cond(1).change_picture_type = 0
 MDIForm1.set_picture_for_change.Enabled = True
 MDIForm1.set_change_type.Enabled = False
 set_change_fig = 0
 set_change_type_ = False
End Sub

Private Sub distance_p_line_Click()
'Call window_fore_rear
operator = "measure"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 3
old_operator = operator
Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2235, "")
'Draw_form.SetFocus
'protect_munu = 1
End Sub

Private Sub draw_circle_tnagent_to_circle_Click()
  Call remove_uncomplete_operat(old_operator)
  list_type_for_draw = 5
   draw_step = 0
    Call init_draw_data
     old_operator = operator
 ' MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4000, "")
  Draw_form.HScroll1.visible = False

End Sub

Private Sub draw_circle_tngent_to_two_circles_Click()
  Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 6
   draw_step = 0
    Call init_draw_data
     old_operator = operator
  'MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4000, "")
  Draw_form.HScroll1.visible = False

End Sub

Private Sub Draw_Click()
  '作图操作
Call clear_wenti_display
If picture_copy = True Then
  'Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, _
          1, chose_w_no%, 0, 0, 0, 2)
 Set C_display_picture = Nothing
 Set C_display_picture = New display_picture
 Call Draw_form.Cls
 Call BitBlt(Draw_form.hdc, 0, 0, Draw_form.Picture1.width, _
     Draw_form.Picture1.Height, Draw_form.Picture1.hdc, 0, 0, vbSrcCopy)
picture_copy = False
End If
If wenti_type > 2 Then
 wenti_type = wenti_type - 2
End If
 End Sub
 
Private Sub e_polygon3_Click()
operator = "epolygon"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 1
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2000, "")
'Draw_form.SetFocus
'protect_munu = 1
End Sub

Private Sub e_polygon4_Click()
operator = "epolygon"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 2
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2010, "")
'Draw_form.SetFocus
'protect_munu = 1
End Sub

Private Sub e_polygon5_Click()
operator = "epolygon"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 3
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2015, "")
'Draw_form.SetFocus
'protect_munu = 1
End Sub

Private Sub e_polygon6_Click()
operator = "epolygon"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 4
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2020, "")
'Draw_form.SetFocus

End Sub

Private Sub Eangle_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & eangle.Caption)
inform.Picture1.Cls
Call set_information_list(eangle_)
End Sub


Private Sub eara_Click()
'Call window_fore_rear
operator = "measure"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 3
old_operator = operator
Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2240, "")
'Draw_form.SetFocus
'protect_munu = 1
End Sub

Private Sub eline_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & eline.Caption)
inform.Picture1.Cls
Call set_information_list(eline_)
End Sub

Private Sub epolygon_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & epolygon.Caption)
inform.Picture1.Cls
Call set_information_list(epolygon_)
End Sub

Private Sub equal_angle_line_Click()
Call menu_item_click("draw_point_and_line", 6, 2315)
End Sub

Private Sub equal_line_Click()
Call menu_item_click("draw_point_and_line", 5, 2200)
End Sub

Private Sub examp_Click()
If run_type = 0 Or run_type > 10 Then
Call clear_wenti_display
'Call init_operat
exam_form.top = 580
exam_form.left = 5780
Load exam_form
exam_form.Show
End If
End Sub


Private Sub exit_Click()
Unload Me
End Sub

Private Sub fd_sx_p_Click()
operator = "change_picture"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 5
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2320, "")
'Draw_form.SetFocus
'protect_munu = 1
End Sub

Private Sub File_Click()
If run_type > 0 And run_type < 11 Then
 Exit Sub
End If
Call clear_wenti_display
'Call init_operat
End Sub

Private Sub Four_point_on_circle_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & four_point_on_circle.Caption)
inform.Picture1.Cls
Call set_information_list(point4_on_circle_)
End Sub

Private Sub Help_Click()
Call clear_wenti_display
Call remove_uncomplete_operat(old_operator)
'Call init_operat
End Sub


Private Sub image_of_function_Click()
 operator = "set_function_data"
 If is_set_function_data = 0 Then
  is_set_function_data = 1
   Call draw_coordianter(Wenti_form.Picture3)
   Call set_menu_for_set_function_data0
 ElseIf is_set_function_data = 1 Then
   Call set_menu_for_set_function_data0
 ElseIf is_set_function_data = 2 Then
   Call set_menu_for_set_function_data1
 Else
   Call recove_set_menu_for_set_function_data
 End If
End Sub

Private Sub index_Click()
Dim file_name As String
file_name = App.path + "\hh " + App.path + "\pmjh3.chm"
Call Shell(file_name, vbMaximizedFocus)
'CommonDialog1.HelpFile = App.Path & "\pmjh3.chm"
'CommonDialog1.HelpCommand = cdlHelpContents
End Sub

Private Sub initial_picture_Click()
Dim i%
If last_conditions.last_cond(1).change_picture_type = line_ Then
' Call draw_change_line(5)
 line_for_change.line_no(0).center(1) = line_for_change.line_no(0).center(0)
 line_for_change.line_no(0).coord(0) = m_poi(line_for_change.line_no(0).poi(0)).data(0).data0.coordinate
 line_for_change.line_no(0).coord(1) = m_poi(line_for_change.line_no(0).poi(1)).data(0).data0.coordinate
 line_for_change.direction = 1
 line_for_change.move.X = 0
 line_for_change.move.Y = 0
 line_for_change.similar_ratio = 1
 line_for_change.rote_angle = 0
 line_for_change.line_no(1) = line_for_change.line_no(0)
 is_first_move = True
ElseIf last_conditions.last_cond(1).change_picture_type = polygon_ Then
 Call draw_polygon(Polygon_for_change.p(0), Polygon_for_change.direction)
 Polygon_for_change.move.X = 0
 Polygon_for_change.move.Y = 0
 Polygon_for_change.p(0).center.X = 0
 Polygon_for_change.p(0).center.Y = 0
 For i% = 0 To Polygon_for_change.p(0).total_v - 1
 Polygon_for_change.p(0).center.X = Polygon_for_change.p(0).center.X + _
    m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.X
 Polygon_for_change.p(0).coord(i%).X = m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.X
 Polygon_for_change.p(0).center.Y = Polygon_for_change.p(0).center.Y + _
    m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.Y
 Polygon_for_change.p(0).coord(i%).Y = m_poi(Polygon_for_change.p(0).v(i%)).data(0).data0.coordinate.Y
 Next i%
 Polygon_for_change.p(0).center.X = Polygon_for_change.p(0).center.X / Polygon_for_change.p(0).total_v
 Polygon_for_change.p(0).center.Y = Polygon_for_change.p(0).center.Y / Polygon_for_change.p(0).total_v
 Polygon_for_change.p(0).coord_center = Polygon_for_change.p(0).center
 Polygon_for_change.similar_ratio = 1
 Polygon_for_change.rote_angle = 0
 is_first_move = True
ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
   Draw_form.Circle (Circle_for_change.c_coord.X, Circle_for_change.c_coord.Y), _
     m_Circ(Circle_for_change.c).data(0).data0.radii, QBColor(12)
   Circle_for_change.move.X = 0
   Circle_for_change.move.Y = 0
   Circle_for_change.radii = m_Circ(Circle_for_change.c).data(0).data0.radii
  is_first_move = True
End If

End Sub

Private Sub Inputcond_Click()
If event_statue = wait_for_input_char Or event_statue = _
    wait_for_modify_char Or event_statue = input_char_again Then
     Wenti_form.Picture1.SetFocus
      Exit Sub
Else
If wenti_type > 2 Then
 wenti_type = wenti_type - 2
End If
'Call init_operat
event_statue = wait_for_input_condition
End If
End Sub

Private Sub jisuan_Click()

End Sub

Private Sub knowl_Click()
Dim foo As Long
CommonDialog1.HelpFile = App.path & "\Pmjh2.hlp"
CommonDialog1.HelpCommand = cdlHelpContents
CommonDialog1.ShowHelp
End Sub

Private Sub length_Click()
'Call window_fore_rear
operator = "measure"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 1
old_operator = operator
Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2230, "")
'Draw_form.SetFocus
'protect_munu = 1
End Sub
Private Sub length_of_segment_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & length_of_segment.Caption)
inform.Picture1.Cls
Call set_information_list(line_value_)
End Sub

Private Sub line_given_length_Click()
Call menu_item_click("draw_point_and_line", 4, 2190) '画定长线段
End Sub

Private Sub MDIForm_Initialize()
Dim i%, j%, k%
Dim A%, date_lim%
Dim ch As String * 1
Dim dir As String * 40
'Dim X As Printer
Dim wenti_name As String * 40
'For Each X In Printers
'If X.Orientation = vbPRORPortrait Then
' Set Printer = X
'  Exit For
'End If
'Next
'Printer.Print LoadResString_(284)
'Printer.EndDoc
Call get_regist
Call Set_inpcond

If MDIForm1.SysInfo1.OSVersion <> 0 Then
 If MDIForm1.SysInfo1.OSPlatform = 1 Then
  system_vision = 1
Else
  system_vision = 2
End If
End If
dir = ""
Call GetSystemDirectory(dir, 40)
'除空格
SystemDirectory = ""
For i% = 1 To 40
ch = Mid$(dir, i%, 1)
A% = Asc(ch)
If A% > 0 Then
SystemDirectory = SystemDirectory + ch
End If
Next i%
SystemDirectory = Trim(SystemDirectory)
dir = ""
Call GetWindowsDirectory(dir, 40)
'除空格
WindowsDirectory = ""
For i% = 1 To 40
ch = Mid$(dir, i%, 1)
A% = Asc(ch)
If A% > 0 Then
WindowsDirectory = WindowsDirectory + ch
End If
Next i%
WindowsDirectory = Trim(WindowsDirectory)
'***************
If SystemDirectory = "" Then
If system_vision = 1 Then
   SystemDirectory = WindowsDirectory + "\system"
Else
   SystemDirectory = WindowsDirectory + "\system32"
End If
End If
protect_file(0) = WindowsDirectory & file_name1 '"\media\Chord.wav"
'If MDIForm1.SysInfo1.OSPlatform = 1 Then
protect_file(1) = App.path & "\My_SysInforLib.dll" 'win2000 留作记录
'ElseIf MDIForm1.SysInfo1.OSPlatform = 2 Then
'protect_file(1) = App.Path + "\My_SysInformLib1.dll"
'End If
protect_file(2) = WindowsDirectory & file_name2 '\Update.sys"
protect_data.computer_id = "0000000"
'protect_data.pass_word_for_teacher = "00000"
'computer_id_ = computer_id
'Timer2.Enabled = True
'screenx& = MDIForm1.Width ' Screen.Width
'screeny& = MDIForm1.Height 'Screen.Height
'For Each X In Printers
'If X.Orientation = vbPRORPortrait Then
 'Set Printer = X
  'Exit For
'End If
'Next
'MDIForm1.Width = screenx& 'Screen.Width 'TwipsPerPixelX
'MDIForm1.Height = screeny& 'Screen.Height 'TwipsPerPixelY
inform_picture_visible = 1
inform_treeview_visible = False
event_statue = -1
reduce_level = 3
reduce_level0 = 3
'Load Draw_form
For i% = 2 To 5
MDIForm1.Toolbar1.Buttons(i%).Image = i% - 1
Next i%
For i% = 7 To 19
MDIForm1.Toolbar1.Buttons(i%).Image = i% - 2
Next i%
MDIForm1.Toolbar1.Buttons(i%).Image = 32
MDIForm1.Toolbar1.Buttons(21).Image = 34
'********************************************************
'打开例题文件
Set C_IO = New IO_class
'Call C_IO.set_exam_list("\example.jak")
'last_record_in_file = 0
'file_record.mark = "DOESSOFT"
'Open App.path & "\example.jak" For Random As #1 Len = Len(wenti_record)
'last_record_in_file = 1
'If LOF(1) > 0 Then
'last_record = 1
'Do While EOF(1) <> True
'Get #1, last_record, wenti_record
'A% = Asc(Mid$(wenti_record.name, 1, 1))
'If A% <> 0 Then
'If Mid$(wenti_record.name, Len(wenti_record.name), 1) = "~" Then
'选择题
'wenti_name = Mid$(wenti_record.name, 1, Len(wenti_record.name) - 1)
'Else
'wenti_name = wenti_record.name
'End If
'ReDim Preserve exam_wenti_name(last_record_in_file) As String * 20
' exam_wenti_name(last_record_in_file) = wenti_name
'  last_record_in_file = last_record_in_file + 1
'End If
' last_record = last_record + 1
'Loop
'End If
'************
'chapter(1).text = "第一章 基本概念"
'chapter(1).no = 200
'chapter(2).text = "  一  直线,射线,线段"
'chapter(2).no = 103
'chapter(3).text = "    1.1  直线"
'chapter(3).no = 101
'chapter(4).text = "    1.2  射线和线段"
'chapter(4).no = 102
'chapter(5).text = "    1.3  线段的比较和度量"
'chapter(5).no = 103
'chapter(6).text = "    1.4  线段的和,差于画法"
'chapter(6).no = 104
'chapter(7).text = "  二  角"
'chapter(7).no = 108
'chapter(8).text = "    1.5  角"
'chapter(8).no = 105
'chapter(9).text = "    1.6  角的比较与度量"
'chapter(9).no = 106
'chapter(10).text = "    1.7  角的和差与画法"
'chapter(10).no = 107
'chapter(11).text = "    1.8  角的分类"
'chapter(11).no = 108
'chapter(12).text = "第二章 相交线,平行线"
'chapter(12).no = 300
'chapter(13).text = "  一  相交线,垂线"
'chapter(13).no = 203
'chapter(14).text = "    2.1  相交线,对顶角"
'chapter(14).no = 202
'chapter(15).text = "    2.2  垂线"
'chapter(15).no = 202
'chapter(16).text = "    2.3  同位角,内错角,同旁内角"
'chapter(16).no = 203
'chapter(17).text = "  二  平行线"
'chapter(17).no = 207
'chapter(18).text = "    2.4  平行线"
'chapter(18).no = 204
'chapter(19).text = "    2.5  平行公理"
'chapter(19).no = 205
'chapter(20).text = "    2.6  平行线的判定"
'chapter(20).no = 206
'chapter(21).text = "    2.7  平行线的性质"
'chapter(21).no = 207
'chapter(22).text = "  三  命题,定理,证明"
'chapter(22).no = 209
'chapter(23).text = "    2.8  命题,定理"
'chapter(23).no = 208
'chapter(24).text = "    2.9  证明"
'chapter(24).no = 209
'chapter(25).text = "第三章 三角形"
'chapter(25).no = 400
'chapter(26).text = "  一  三角形"
'chapter(26).no = 303
'chapter(27).text = "   3.1  关于三角形的一些概念"
'chapter(27).no = 301
'chapter(28).text = "   3.2  三角形三条边的关系"
'chapter(28).no = 302
'chapter(29).text = "   3.3  三角形的内角和"
'chapter(29).no = 303
'chapter(30).text = "  二  全等三角形"
'chapter(30).no = 309
'chapter(31).text = "   3.4  三角形全等的判定"
'chapter(31).no = 304
'chapter(32).text = "   3.5  三角形全等的判定"
'chapter(32).no = 305
'chapter(33).text = "   3.6  三角形全等的判定"
'chapter(33).no = 306
'chapter(34).text = "   3.7  直角三角形全等的判定"
'chapter(34).no = 307
'chapter(35).text = "   3.8  角的平分线"
'chapter(35).no = 308
'chapter(36).text = "   3.9  角的平分线"
'chapter(36).no = 309
'chapter(37).text = "  三  尺规作图"
'chapter(37).no = 311
'chapter(38).text = "   3.10  基本作图"
'chapter(38).no = 310
'chapter(39).text = "   3.11  作图题举例"
'chapter(39).no = 311
'chapter(40).text = "  四  等腰三角形"
'chapter(40).no = 315
'chapter(41).text = "   3.12  等腰三角形的性质"
'chapter(41).no = 312
'chapter(42).text = "   3.13  等腰三角形的判定"
'chapter(42).no = 313
'chapter(43).text = "   3.14  线段的垂直平分线"
'chapter(43).no = 314
'chapter(44).text = "   3.15  轴对称和轴对称图形"
'\chapter(44).no = 315
'chapter(45).text = "  五  勾股定理"
'chapter(45).no = 317
'chapter(46).text = "   3.16  勾股定理"
'chapter(46).no = 316
'chapter(47).text = "   3.17  勾股定理的逆定理"
'chapter(47).no = 317
'chapter(48).text = "第四章  四边形"
'chapter(48).no = 500
'chapter(49).text = "  一  四边形"
'chapter(49).no = 402
'chapter(50).text = "   4.1  四边形"
'chapter(50).no = 401
'chapter(51).text = "   4.2  多边形的内角和"
'chapter(51).no = 402
'chapter(52).text = "  二  平行四边形"
'chapter(52).no = 407
'chapter(53).text = "   4.3  平行四边形及其性质"
'chapter(53).no = 403
'chapter(54).text = "    4.4  平行四边形的判定"
'chapter(54).no = 404
'chapter(55).text = "    4.5  矩形,菱性"
'chapter(55).no = 405
'chapter(56).text = "    4.6  正方形"
'chapter(56).no = 406
'chapter(57).text = "    4.7  中心对称和中心对称图形"
'chapter(57).no = 407
'chapter(58).text = "  三  梯形"
'chapter(58).no = 410
'chapter(59).text = "    4.8  梯形"
'chapter(59).no = 408
'chapter(60).text = "    4.9  平行线分线段定理"
'chapter(60).no = 409
'chapter(61).text = "    4.10  三角形,梯形的中位线"
'chapter(61).no = 410
'chapter(62).text = "第五章  相似形"
'chapter(62).no = 600
'chapter(63).text = "  一  比例线段"
'chapter(63).no = 502
'chapter(64).text = "    5.1  比例线段"
'chapter(64).no = 501
'chapter(65).text = "    5.2 平行线分线段成比例定理"
'chapter(65).no = 502
'chapter(66).text = "  二  相似三角形"
'chapter(66).no = 506
'chapter(67).text = "    5.3  相似三角形"
'chapter(67).no = 503
'chapter(68).text = "    5.4  三角形相似的判定"
'chapter(68).no = 504
'chapter(69).text = "    5.5  三角形相似的性质"
'chapter(69).no = 505
'chapter(70).text = "    5.6  相似多边形"
'chapter(70).no = 506
'chapter(71).text = "第六章  解直角三角形"
'chapter(71).no = 602
'chapter(72).text = "  一  锐角三角函数"
'chapter(72).no = 303
'chapter(73).text = "    6.1  正弦和余弦"
'chapter(73).no = 601
'chapter(74).text = "    6.2  正切和余切"
'chapter(74).no = 602
'chapter(75).text = "  二  解直角三角形"
'chapter(75).no = 605
'chapter(76).text = "    6.3  解直角三角形"
'chapter(76).no = 603
'chapter(77).text = "    6.4  应用举例"
'chapter(77).no = 604
'chapter(78).text = "    6.5  实习作业"
'chapter(78).no = 605
'chapter(79).text = "第七章  圆"
'chapter(79).no = 303
'chapter(80).text = "  一  圆的有关性质"
'chapter(80).no = 800
'chapter(81).text = "    7.1  圆"
'chapter(81).no = 706
'chapter(82).text = "    7.2  过三点的圆"
'chapter(82).no = 702
'chapter(83).text = "    7.3  垂直于弦的直径"
'chapter(83).no = 703
'chapter(84).text = "    7.4  圆心角,弧,弦,弦心距之间的关系"
'chapter(84).no = 704
'chapter(85).text = "    7.5  圆周角"
'chapter(85).no = 705
'chapter(86).text = "    7.6  圆的内接四边形"
'chapter(86).no = 706
'chapter(87).text = "  二  直线和圆的位置关系"
'chapter(87).no = 712
'chapter(88).text = "    7.7  直线和圆的位置关系"
'chapter(88).no = 707
'chapter(89).text = "    7.8  切线的判定和性质"
'chapter(89).no = 708
'chapter(90).text = "    7.9  三角形的内切圆"
'chapter(90).no = 709
'chapter(91).text = "   *7.10  切线长定理"
'chapter(91).no = 710
'chapter(92).text = "   *7.11  弦切角"
'chapter(92).no = 711
'chapter(93).text = "   *7.12  和圆有关的比例线段"
'chapter(93).no = 712
'chapter(94).text = "  三  圆和圆的位置关系"
'chapter(94).no = 715
'chapter(95).text = "    7.13  圆和圆的位置关系"
'chapter(95).no = 713
'chapter(96).text = "    7.14  两圆的公切段"
'chapter(96).no = 714
'chapter(97).text = "    7.15  相切在作图中的应用"
'chapter(97).no = 715
'chapter(98).text = "  四  正多边形和圆"
'chapter(98).no = 721
'chapter(99).text = "    7.16  正多边形和圆"
'chapter(99).no = 716
'chapter(100).text = "    7.17  正多边形的有关计算"
'chapter(100).no = 717
'chapter(101).text = "    7.18  画正多边形"
'chapter(101).no = 718
'chapter(102).text = "    7.19  圆周长,弧长"
'chapter(102).no = 719
'chapter(103).text = "    7.20  圆,扇形,弓形的面积"
'chapter(103).no = 720
'chapter(104).text = "    7.21  圆柱和圆锥的侧面展开面"
'chapter(104).no = 721
'chapter(105).text = "补充定理"
'chapter(105).no = 900
'last_chapter = 106
'chapter_no = 1000
th_chose(-6).chose = 1
th_chose(-6).TH_name = LoadResString_(1755, "") '"塞瓦定理"
'th_chose(-6).chose = 0
th_chose(-6).text = LoadResString_(1760, "") '"塞瓦定理:过A、B、C连接" + LoadResString_(875) + "ABC内的一点P的直线分别交于DEF,则(AD/DC)(CE/ED)(DF/FA)=1." 'LoadResString_(375)
th_chose(-5).chose = 1
th_chose(-5).TH_name = LoadResString_(1940, "")
'th_chose(-5).chose = 0
'th_chose(-5).text = LoadResString_(375)
th_chose(-4).TH_name = LoadResString_(1945, "")
th_chose(-4).chose = 1
th_chose(-4).text = LoadResString_(1355, "") '"直线与" + LoadResString_(875) + "ABC的三边或其延长线分别交于DEF,则(AD/DC)(CE/ED)(DF/FA)=1." 'LoadResString_(375)
th_chose(-4).chapter = 0
th_chose(-3).TH_name = LoadResString_(1950, "")
th_chose(-3).chose = 1
th_chose(-3).text = LoadResString_(1950, "")
th_chose(-3).chapter = 0
th_chose(-2).TH_name = LoadResString_(2070, "")
th_chose(-2).chose = 1
th_chose(-2).text = LoadResString_(2075, "")
th_chose(-2).chapter = 0
th_chose(-1).TH_name = LoadResString_(2080, "")
th_chose(-1).chose = 1
th_chose(-1).text = LoadResString_(2085, "")
th_chose(-1).chapter = 0
th_chose(1).TH_name = LoadResString_(2090, "")
th_chose(1).chose = 1
th_chose(1).text = LoadResString_(2095, "")
th_chose(1).chapter = 0
th_chose(2).TH_name = LoadResString_(2100, "")
th_chose(2).chose = 1
th_chose(2).text = LoadResString_(2330, "")
th_chose(2).chapter = 0
th_chose(3).TH_name = LoadResString_(2335, "")
th_chose(3).chose = 1
th_chose(3).text = LoadResString_(2340, "")
th_chose(3).chapter = 0
th_chose(4).TH_name = LoadResString_(2345, "")
th_chose(4).chose = 1
th_chose(4).text = LoadResString_(2350, "")
th_chose(4).chapter = 0
th_chose(5).TH_name = LoadResString_(2355, "")
th_chose(5).chose = 1
th_chose(5).text = LoadResString_(2360, "")
th_chose(5).chapter = 0
th_chose(6).TH_name = LoadResString_(2365, "")
th_chose(6).chose = 1
th_chose(6).text = LoadResString_(2365, "")
th_chose(6).chapter = 201
th_chose(7).TH_name = LoadResString_(2370, "")
th_chose(7).chose = 1
th_chose(7).text = LoadResString_(2375, "")
th_chose(7).chapter = 204
th_chose(8).TH_name = LoadResString_(2380, "")
th_chose(8).chose = 1
th_chose(8).text = LoadResString_(2385, "")
th_chose(8).chapter = 206
th_chose(9).TH_name = LoadResString_(2390, "")
th_chose(9).chose = 1
th_chose(9).text = LoadResString_(2395, "")
th_chose(9).chapter = 206
th_chose(10).TH_name = LoadResString_(2400, "")
th_chose(10).chose = 1
th_chose(10).text = LoadResString_(2405, "")
th_chose(10).chapter = 206
th_chose(11).TH_name = LoadResString_(2410, "")
th_chose(11).chose = 1
th_chose(11).text = LoadResString_(2415, "")
th_chose(11).chapter = 207
th_chose(12).TH_name = LoadResString_(2420, "")
th_chose(12).chose = 1
th_chose(12).text = LoadResString_(2425, "")
th_chose(12).chapter = 207
th_chose(13).TH_name = LoadResString_(2430, "")
th_chose(13).chose = 1
th_chose(13).text = LoadResString_(2435, "")
th_chose(13).chapter = 207
th_chose(14).TH_name = LoadResString_(2440, "")
th_chose(14).chose = 1
th_chose(14).text = LoadResString_(2445, "")
th_chose(14).chapter = 207
th_chose(15).TH_name = LoadResString_(2450, "")
th_chose(15).chose = 1
th_chose(15).text = LoadResString_(2455, "")
th_chose(15).chapter = 208
th_chose(16).TH_name = LoadResString_(2460, "")
th_chose(16).chose = 1
th_chose(16).text = LoadResString_(2465, "")
th_chose(16).chapter = 208
th_chose(17).TH_name = LoadResString_(2470, "")
th_chose(17).chose = 1
th_chose(17).text = LoadResString_(2475, "")
th_chose(17).chapter = 208
th_chose(18).TH_name = LoadResString_(2480, "")
th_chose(18).chose = 1
th_chose(18).text = LoadResString_(2485, "")
th_chose(18).chapter = 302
th_chose(19).TH_name = LoadResString_(2490, "")
th_chose(19).chose = 1
th_chose(19).text = LoadResString_(2495, "")
th_chose(19).chapter = 302
th_chose(20).TH_name = LoadResString_(2500, "")
th_chose(20).chose = 1
th_chose(20).text = LoadResString_(2505, "")
th_chose(20).chapter = 303
th_chose(21).TH_name = LoadResString_(2510, "")
th_chose(21).chose = 1
th_chose(21).text = LoadResString_(2515, "")
th_chose(21).chapter = 303
th_chose(22).TH_name = LoadResString_(2520, "")
th_chose(22).chose = 1
th_chose(22).text = LoadResString_(2525, "")
th_chose(22).chapter = 303
th_chose(23).TH_name = LoadResString_(2530, "")
th_chose(23).chose = 1
th_chose(23).text = LoadResString_(2535, "")
th_chose(23).chapter = 303
th_chose(24).TH_name = LoadResString_(2540, "")
th_chose(24).chose = 1
th_chose(24).text = LoadResString_(2545, "")
th_chose(24).chapter = 304
th_chose(25).TH_name = LoadResString_(2550, "")
th_chose(25).chose = 1
th_chose(25).text = LoadResString_(2555, "")
th_chose(25).chapter = 303
th_chose(26).TH_name = LoadResString_(2560, "")
th_chose(26).chose = 1
th_chose(26).text = LoadResString_(2565, "")
th_chose(26).chapter = 303
th_chose(27).TH_name = LoadResString_(2570, "")
th_chose(27).chose = 1
th_chose(27).text = LoadResString_(2575, "")
th_chose(27).chapter = 303
th_chose(28).TH_name = LoadResString_(2580, "")
th_chose(28).chose = 1
th_chose(28).text = LoadResString_(2585, "")
th_chose(28).chapter = 304
th_chose(29).TH_name = LoadResString_(2590, "")
th_chose(29).chose = 1
th_chose(29).text = LoadResString_(2595, "")
th_chose(29).chapter = 304
th_chose(30).TH_name = LoadResString_(2600, "")
th_chose(30).chose = 1
th_chose(30).text = LoadResString_(2605, "")
th_chose(30).chapter = 304
th_chose(31).TH_name = LoadResString_(2610, "")
th_chose(31).chose = 1
th_chose(31).text = LoadResString_(2615, "")
th_chose(31).chapter = 304
th_chose(32).TH_name = LoadResString_(2620, "")
th_chose(32).chose = 1
th_chose(32).text = LoadResString_(2625, "")
th_chose(32).chapter = 304
th_chose(33).TH_name = LoadResString_(2630, "")
th_chose(33).chose = 1
th_chose(33).text = LoadResString_(2635, "")
th_chose(33).chapter = 309
th_chose(34).TH_name = LoadResString_(2640, "")
th_chose(34).chose = 1
th_chose(34).text = LoadResString_(2645, "")
th_chose(34).chapter = 309
th_chose(35).TH_name = LoadResString_(2650, "")
th_chose(35).chose = 1
th_chose(35).text = LoadResString_(2655, "")
th_chose(35).chapter = 309
th_chose(36).TH_name = LoadResString_(2660, "")
th_chose(36).chose = 1
th_chose(36).text = LoadResString_(2665, "")
th_chose(36).chapter = 312
th_chose(37).TH_name = LoadResString_(2670, "")
th_chose(37).chose = 1
th_chose(37).text = LoadResString_(2675, "")
th_chose(37).chapter = 312
th_chose(38).TH_name = LoadResString_(2680, "")
th_chose(38).chose = 1
th_chose(38).text = LoadResString_(2685, "")
th_chose(38).chapter = 312
th_chose(39).TH_name = LoadResString_(2690, "")
th_chose(39).chose = 1
th_chose(39).text = LoadResString_(2695, "")
th_chose(39).chapter = 312
th_chose(40).TH_name = LoadResString_(2700, "")
th_chose(40).chose = 1
th_chose(40).text = LoadResString_(2705, "")
th_chose(40).chapter = 313
th_chose(41).TH_name = LoadResString_(2710, "")
th_chose(41).chose = 1
th_chose(41).text = LoadResString_(2715, "")
th_chose(41).chapter = 313
th_chose(42).TH_name = LoadResString_(2720, "")
th_chose(42).chose = 1
th_chose(42).text = LoadResString_(2725, "")
th_chose(42).chapter = 313
th_chose(43).TH_name = LoadResString_(2730, "")
th_chose(43).chose = 1
th_chose(43).text = LoadResString_(2735, "")
th_chose(43).chapter = 313
th_chose(44).TH_name = LoadResString_(2740, "")
th_chose(44).chose = 1
th_chose(44).text = LoadResString_(2745, "")
th_chose(44).chapter = 314
th_chose(45).TH_name = LoadResString_(2750, "")
th_chose(45).chose = 1
th_chose(45).text = LoadResString_(2755, "")
th_chose(45).chapter = 314
th_chose(46).TH_name = LoadResString_(2760, "")
th_chose(46).chose = 1
th_chose(46).text = LoadResString_(2765, "")
th_chose(46).chapter = 314
th_chose(47).TH_name = LoadResString_(2770, "")
th_chose(47).chose = 1
th_chose(47).text = LoadResString_(2775, "")
th_chose(47).chapter = 315
th_chose(48).TH_name = LoadResString_(2780, "")
th_chose(48).chose = 1
th_chose(48).text = LoadResString_(2785, "")
th_chose(48).chapter = 315
th_chose(49).TH_name = LoadResString_(2790, "")
th_chose(49).chose = 1
th_chose(49).text = LoadResString_(2795, "")
th_chose(49).chapter = 315
th_chose(50).TH_name = LoadResString_(2800, "")
th_chose(50).chose = 1
th_chose(50).text = LoadResString_(2805, "")
th_chose(50).chapter = 315
th_chose(51).TH_name = LoadResString_(2810, "")
th_chose(51).chose = 1
th_chose(51).text = LoadResString_(2815, "")
th_chose(51).chapter = 316
th_chose(52).TH_name = LoadResString_(2100, "")
th_chose(52).chose = 1
th_chose(52).text = LoadResString_(2825, "")
th_chose(52).chapter = 316
th_chose(53).TH_name = LoadResString_(2830, "")
th_chose(53).chose = 1
th_chose(53).text = LoadResString_(2835, "")
th_chose(53).chapter = 401
th_chose(54).TH_name = LoadResString_(2840, "")
th_chose(54).chose = 1
th_chose(54).text = LoadResString_(2845, "")
th_chose(54).chapter = 401
th_chose(55).TH_name = LoadResString_(2850, "")
th_chose(55).chose = 1
th_chose(55).text = LoadResString_(2855, "")
th_chose(55).chapter = 402
th_chose(56).TH_name = LoadResString_(2860, "")
th_chose(56).chose = 1
th_chose(56).text = LoadResString_(2865, "")
th_chose(56).chapter = 402
th_chose(57).TH_name = LoadResString_(2870, "")
th_chose(57).chose = 1
th_chose(57).text = LoadResString_(2875, "")
th_chose(57).chapter = 403
th_chose(58).TH_name = LoadResString_(2880, "")
th_chose(58).chose = 1
th_chose(58).text = LoadResString_(2885, "")
th_chose(58).chapter = 403
th_chose(59).TH_name = LoadResString_(2890, "")
th_chose(59).chose = 1
th_chose(59).text = LoadResString_(2895, "")
th_chose(59).chapter = 403
th_chose(60).TH_name = LoadResString_(2900, "")
th_chose(60).chose = 1
th_chose(60).text = LoadResString_(2905, "")
th_chose(60).chapter = 403
th_chose(61).TH_name = LoadResString_(2910, "")
th_chose(61).chose = 1
th_chose(61).text = LoadResString_(2915, "")
th_chose(61).chapter = 403
th_chose(62).TH_name = LoadResString_(2920, "")
th_chose(62).chose = 1
th_chose(62).text = LoadResString_(2925, "")
th_chose(62).chapter = 404
th_chose(63).TH_name = LoadResString_(2930, "")
th_chose(63).chose = 1
th_chose(63).text = LoadResString_(2935, "")
th_chose(63).chapter = 404
th_chose(64).TH_name = LoadResString_(2940, "")
th_chose(64).chose = 1
th_chose(64).text = LoadResString_(2945, "")
th_chose(64).chapter = 404
th_chose(65).TH_name = LoadResString_(2950, "")
th_chose(65).chose = 1
th_chose(65).text = LoadResString_(2955, "")
th_chose(65).chapter = 404
th_chose(66).TH_name = LoadResString_(2960, "")
th_chose(66).chose = 1
th_chose(66).text = LoadResString_(2965, "")
th_chose(66).chapter = 404
th_chose(67).TH_name = LoadResString_(2970, "")
th_chose(67).chose = 1
th_chose(67).text = LoadResString_(2975, "")
th_chose(67).chapter = 405
th_chose(68).TH_name = LoadResString_(2980, "")
th_chose(68).chose = 1
th_chose(68).text = LoadResString_(2985, "")
th_chose(68).chapter = 405
th_chose(69).TH_name = LoadResString_(2990, "")
th_chose(69).chose = 1
th_chose(69).text = LoadResString_(2995, "")
th_chose(69).chapter = 405
th_chose(70).TH_name = LoadResString_(3000, "")
th_chose(70).chose = 1
th_chose(70).text = LoadResString_(3005, "")
th_chose(70).chapter = 405
th_chose(71).TH_name = LoadResString_(3010, "")
th_chose(71).chose = 1
th_chose(71).text = LoadResString_(3015, "")
th_chose(71).chapter = 405
th_chose(72).TH_name = LoadResString_(3020, "")
th_chose(72).chose = 1
th_chose(72).text = LoadResString_(3025, "")
th_chose(72).chapter = 405
th_chose(73).TH_name = LoadResString_(3030, "")
th_chose(73).chose = 1
th_chose(73).text = LoadResString_(3035, "")
th_chose(73).chapter = 405
th_chose(74).TH_name = LoadResString_(3040, "")
th_chose(74).chose = 1
th_chose(74).text = LoadResString_(3045, "")
th_chose(74).chapter = 405
th_chose(75).TH_name = LoadResString_(3050, "")
th_chose(75).chose = 1
th_chose(75).text = LoadResString_(3055, "")
th_chose(75).chapter = 405
th_chose(76).TH_name = LoadResString_(3060, "")
th_chose(76).chose = 1
th_chose(76).text = LoadResString_(3065, "")
th_chose(76).chapter = 405
th_chose(77).TH_name = LoadResString_(3070, "")
th_chose(77).chose = 1
th_chose(77).text = LoadResString_(3075, "")
th_chose(77).chapter = 405
th_chose(78).TH_name = LoadResString_(3080, "")
th_chose(78).chose = 1
th_chose(78).text = LoadResString_(3085, "")
th_chose(78).chapter = 405
th_chose(79).TH_name = LoadResString_(3090, "")
th_chose(79).chose = 1
th_chose(79).text = LoadResString_(3095, "")
th_chose(79).chapter = 406
th_chose(80).TH_name = LoadResString_(3100, "")
th_chose(80).chose = 1
th_chose(80).text = LoadResString_(3105, "")
th_chose(80).chapter = 406
th_chose(81).TH_name = LoadResString_(3110, "")
th_chose(81).chose = 1
th_chose(81).text = LoadResString_(3115, "")
th_chose(81).chapter = 406
th_chose(82).TH_name = LoadResString_(3120, "")
th_chose(82).chose = 1
th_chose(82).text = LoadResString_(3125, "")
th_chose(82).chapter = 406
th_chose(83).TH_name = LoadResString_(3130, "")
th_chose(83).chose = 1
th_chose(83).text = LoadResString_(3135, "")
th_chose(83).chapter = 406
th_chose(84).TH_name = LoadResString_(3140, "")
th_chose(84).chose = 1
th_chose(84).text = LoadResString_(3145, "")
th_chose(84).chapter = 406
th_chose(85).TH_name = LoadResString_(3150, "")
th_chose(85).chose = 1
th_chose(85).text = LoadResString_(3155, "")
th_chose(85).chapter = 406
th_chose(86).TH_name = LoadResString_(3160, "")
th_chose(86).chose = 1
th_chose(86).text = LoadResString_(3165, "")
th_chose(86).chapter = 407
th_chose(87).TH_name = LoadResString_(3170, "")
th_chose(87).chose = 1
th_chose(87).text = LoadResString_(3175, "")
th_chose(87).chapter = 407
th_chose(88).TH_name = LoadResString_(3180, "")
th_chose(88).chose = 1
th_chose(88).text = LoadResString_(3185, "")
th_chose(88).chapter = 407
th_chose(89).TH_name = LoadResString_(3190, "")
th_chose(89).chose = 1
th_chose(89).text = LoadResString_(3195, "")
th_chose(89).chapter = 408
th_chose(90).TH_name = LoadResString_(3200, "")
th_chose(90).chose = 1
th_chose(90).text = LoadResString_(3205, "")
th_chose(90).chapter = 408
th_chose(91).TH_name = LoadResString_(3210, "")
th_chose(91).chose = 1
th_chose(91).text = LoadResString_(3215, "")
th_chose(91).chapter = 408
th_chose(92).TH_name = LoadResString_(3220, "")
th_chose(92).chose = 1
th_chose(92).text = LoadResString_(3225, "")
th_chose(92).chapter = 408
th_chose(93).TH_name = LoadResString_(3230, "")
th_chose(93).chose = 1
th_chose(93).text = LoadResString_(3235, "")
th_chose(93).chapter = 408
th_chose(94).TH_name = LoadResString_(3240, "")
th_chose(94).chose = 1
th_chose(94).text = LoadResString_(3245, "")
th_chose(94).chapter = 409
th_chose(95).TH_name = LoadResString_(3250, "")
th_chose(95).chose = 1
th_chose(95).text = LoadResString_(3255, "")
th_chose(95).chapter = 409
th_chose(96).TH_name = LoadResString_(3260, "")
th_chose(96).chose = 1
th_chose(96).text = LoadResString_(3265, "")
th_chose(96).chapter = 409
th_chose(97).TH_name = LoadResString_(3270, "")
th_chose(97).chose = 1
th_chose(97).text = LoadResString_(3275, "")
th_chose(97).chapter = 410
th_chose(98).TH_name = LoadResString_(3280, "")
th_chose(98).chose = 1
th_chose(98).text = LoadResString_(3285, "")
th_chose(98).chapter = 410
th_chose(99).TH_name = LoadResString_(3290, "")
th_chose(99).chose = 1
th_chose(99).text = LoadResString_(3295, "")
th_chose(99).chapter = 502
th_chose(100).TH_name = LoadResString_(3300, "")
th_chose(100).chose = 1
th_chose(100).text = LoadResString_(3305, "")
th_chose(100).chapter = 502
th_chose(101).TH_name = LoadResString_(3310, "")
th_chose(101).chose = 1
th_chose(101).text = LoadResString_(3315, "")
th_chose(101).chapter = 502
th_chose(102).TH_name = LoadResString_(3320, "")
th_chose(102).chose = 1
th_chose(102).text = LoadResString_(3325, "")
th_chose(102).chapter = 503
th_chose(103).TH_name = LoadResString_(3330, "")
th_chose(103).chose = 1
th_chose(103).text = LoadResString_(3335, "")
th_chose(103).chapter = 503
th_chose(104).TH_name = LoadResString_(3340, "")
th_chose(104).chose = 1
th_chose(104).text = LoadResString_(3345, "")
th_chose(104).chapter = 503
th_chose(105).TH_name = LoadResString_(3350, "")
th_chose(105).chose = 1
th_chose(105).text = LoadResString_(3355, "")
th_chose(105).chapter = 504
th_chose(106).TH_name = LoadResString_(3360, "")
th_chose(106).chose = 1
th_chose(106).text = LoadResString_(3365, "")
th_chose(106).chapter = 504
th_chose(107).TH_name = LoadResString_(3370, "")
th_chose(107).chose = 1
th_chose(107).text = LoadResString_(3375, "")
th_chose(107).chapter = 504
th_chose(108).TH_name = LoadResString_(3380, "")
th_chose(108).chose = 1
th_chose(108).text = LoadResString_(3385, "")
th_chose(108).chapter = 504
th_chose(109).TH_name = LoadResString_(3390, "")
th_chose(109).chose = 1
th_chose(109).text = LoadResString_(3395, "")
th_chose(109).chapter = 505
th_chose(110).TH_name = LoadResString_(3400, "")
th_chose(110).chose = 1
th_chose(110).text = LoadResString_(3405, "")
th_chose(110).chapter = 505
th_chose(111).TH_name = LoadResString_(3410, "")
th_chose(111).chose = 1
th_chose(111).text = LoadResString_(3415, "")
th_chose(111).chapter = 505
th_chose(112).TH_name = LoadResString_(3420, "")
th_chose(112).chose = 1
th_chose(112).text = LoadResString_(3425, "")
th_chose(112).chapter = 505
th_chose(113).TH_name = LoadResString_(3430, "")
th_chose(113).chose = 1
th_chose(113).text = LoadResString_(3435, "")
th_chose(113).chapter = 506
th_chose(114).TH_name = LoadResString_(3440, "")
th_chose(114).chose = 1
th_chose(114).text = LoadResString_(3445, "")
th_chose(114).chapter = 506
th_chose(115).TH_name = LoadResString_(3450, "")
th_chose(115).chose = 1
th_chose(115).text = LoadResString_(3455, "")
th_chose(115).chapter = 506
th_chose(116).TH_name = LoadResString_(3460, "")
th_chose(116).chose = 1
th_chose(116).text = LoadResString_(3465, "")
th_chose(116).chapter = 506
th_chose(117).TH_name = LoadResString_(3470, "")
th_chose(117).chose = 1
th_chose(117).text = LoadResString_(3475, "")
th_chose(117).chapter = 506
th_chose(118).TH_name = LoadResString_(3480, "")
th_chose(118).chose = 1
th_chose(118).text = LoadResString_(3485, "")
th_chose(118).chapter = 600
th_chose(119).TH_name = LoadResString_(3490, "")
th_chose(119).chose = 1
th_chose(119).text = LoadResString_(3495, "")
th_chose(119).chapter = 702
th_chose(120).TH_name = LoadResString_(3500, "")
th_chose(120).chose = 1
th_chose(120).text = LoadResString_(3505, "")
th_chose(120).chapter = 703
th_chose(121).TH_name = LoadResString_(3510, "")
th_chose(121).chose = 1
th_chose(121).text = LoadResString_(3515, "")
th_chose(121).chapter = 703
th_chose(122).TH_name = LoadResString_(3520, "")
th_chose(122).chose = 1
th_chose(122).text = LoadResString_(3525, "")
th_chose(122).chapter = 703
th_chose(123).TH_name = LoadResString_(3530, "")
th_chose(123).chose = 1
th_chose(123).text = LoadResString_(3535, "")
th_chose(123).chapter = 703
th_chose(124).TH_name = LoadResString_(3540, "")
th_chose(124).chose = 1
th_chose(124).text = LoadResString_(3545, "")
th_chose(124).chapter = 703
th_chose(125).TH_name = LoadResString_(3550, "")
th_chose(125).chose = 1
th_chose(125).text = LoadResString_(3555, "")
th_chose(125).chapter = 704
th_chose(126).TH_name = LoadResString_(3560, "")
th_chose(126).chose = 1
th_chose(126).text = LoadResString_(3565, "")
th_chose(126).chapter = 704
th_chose(127).TH_name = LoadResString_(3570, "")
th_chose(127).chose = 1
th_chose(127).text = LoadResString_(3575, "")
th_chose(127).chapter = 705
th_chose(128).TH_name = LoadResString_(3580, "")
th_chose(128).chose = 1
th_chose(128).text = LoadResString_(3585, "")
th_chose(128).chapter = 705
th_chose(129).TH_name = LoadResString_(3590, "")
th_chose(129).chose = 1
th_chose(129).text = LoadResString_(3595, "")
th_chose(129).chapter = 705
th_chose(130).TH_name = LoadResString_(3600, "")
th_chose(130).chose = 1
th_chose(130).text = LoadResString_(3605, "")
th_chose(130).chapter = 705
th_chose(131).TH_name = LoadResString_(3610, "")
th_chose(131).chose = 1
th_chose(131).text = LoadResString_(3615, "")
th_chose(131).chapter = 706
th_chose(132).TH_name = LoadResString_(3620, "")
th_chose(132).chose = 1
th_chose(132).text = LoadResString_(3625, "")
th_chose(132).chapter = 706
th_chose(133).TH_name = LoadResString_(3630, "")
th_chose(133).chose = 1
th_chose(133).text = LoadResString_(3635, "")
th_chose(133).chapter = 706
th_chose(134).TH_name = LoadResString_(3640, "")
th_chose(134).chose = 1
th_chose(134).text = LoadResString_(3645, "")
th_chose(134).chapter = 706
th_chose(135).TH_name = LoadResString_(3650, "")
th_chose(135).chose = 1
th_chose(135).text = LoadResString_(3655, "")
th_chose(135).chapter = 706
th_chose(136).TH_name = LoadResString_(3660, "")
th_chose(136).chose = 1
th_chose(136).text = LoadResString_(3665, "")
th_chose(136).chapter = 708
th_chose(137).TH_name = LoadResString_(3670, "")
th_chose(137).chose = 1
th_chose(137).text = LoadResString_(3675, "")
th_chose(137).chapter = 708
th_chose(138).TH_name = LoadResString_(3680, "")
th_chose(138).chose = 1
th_chose(138).text = LoadResString_(3685, "")
th_chose(138).chapter = 708
th_chose(139).TH_name = LoadResString_(3690, "")
th_chose(139).chose = 1
th_chose(139).text = LoadResString_(3695, "")
th_chose(139).chapter = 708
th_chose(140).TH_name = LoadResString_(3700, "")
th_chose(140).chose = 1
th_chose(140).text = LoadResString_(3705, "")
th_chose(140).chapter = 710
th_chose(141).TH_name = LoadResString_(3710, "")
th_chose(141).chose = 1
th_chose(141).text = LoadResString_(3715, "")
th_chose(141).chapter = 711
th_chose(142).TH_name = LoadResString_(3720, "")
th_chose(142).chose = 1
th_chose(142).text = LoadResString_(3725, "")
th_chose(142).chapter = 711
th_chose(143).TH_name = LoadResString_(3730, "")
th_chose(143).chose = 1
th_chose(143).text = LoadResString_(3735, "")
th_chose(143).chapter = 712
th_chose(144).TH_name = LoadResString_(3740, "")
th_chose(144).chose = 1
th_chose(144).text = LoadResString_(3745, "")
th_chose(144).chapter = 712
th_chose(145).TH_name = LoadResString_(3750, "")
th_chose(145).chose = 1
th_chose(145).text = LoadResString_(3755, "")
th_chose(145).chapter = 712
th_chose(146).TH_name = LoadResString_(3760, "")
th_chose(146).chose = 1
th_chose(146).text = LoadResString_(3765, "")
th_chose(146).chapter = 712
th_chose(147).TH_name = LoadResString_(3770, "")
th_chose(147).chose = 1
th_chose(147).text = LoadResString_(3775, "")
th_chose(147).chapter = 713
th_chose(148).TH_name = LoadResString_(3780, "")
th_chose(148).chose = 1
th_chose(148).text = LoadResString_(3785, "")
th_chose(148).chapter = 713
th_chose(149).TH_name = LoadResString_(3790, "")
th_chose(149).chose = 1
th_chose(149).text = LoadResString_(3795, "")
th_chose(149).chapter = 714
th_chose(150).TH_name = LoadResString_(3800, "")
th_chose(150).chose = 1
th_chose(150).text = LoadResString_(3805, "")
th_chose(150).chapter = 714
th_chose(151).TH_name = LoadResString_(3810, "")
th_chose(151).chose = 1
th_chose(151).text = LoadResString_(3815, "")
th_chose(151).chapter = 714
th_chose(152).TH_name = LoadResString_(3820, "")
th_chose(152).chose = 1
th_chose(152).text = LoadResString_(3825, "")
th_chose(152).chapter = 714
th_chose(153).TH_name = LoadResString_(3830, "")
th_chose(153).chose = 0
th_chose(153).text = LoadResString_(3835, "")
th_chose(153).chapter = 800
th_chose(154).TH_name = LoadResString_(3840, "")
th_chose(154).chose = 0
th_chose(154).text = LoadResString_(3845, "")
th_chose(154).chapter = 800
th_chose(155).TH_name = LoadResString_(3850, "")
th_chose(155).chose = 0
th_chose(155).text = LoadResString_(3855, "")
th_chose(155).chapter = 800
th_chose(156).TH_name = LoadResString_(3860, "")
th_chose(156).chose = 0
th_chose(156).text = LoadResString_(3865, "")
th_chose(156).chapter = 800
th_chose(157).TH_name = LoadResString_(3870, "")
th_chose(157).chose = 0
th_chose(157).text = LoadResString_(3875, "")
th_chose(157).chapter = 800
th_chose(158).TH_name = LoadResString_(3880, "")
th_chose(158).chose = 0
th_chose(158).text = LoadResString_(3885, "")
th_chose(158).chapter = 800
th_chose(159).TH_name = LoadResString_(3890, "")
th_chose(159).chose = 1
th_chose(159).text = LoadResString_(3895, "")
th_chose(159).chapter = 800
th_chose(160).TH_name = LoadResString_(3900, "")
th_chose(160).chose = 1
th_chose(160).text = LoadResString_(3905, "")
th_chose(160).chapter = 800
th_chose(161).TH_name = LoadResString_(3910, "")
th_chose(161).chose = 1
th_chose(161).text = LoadResString_(3915, "")
th_chose(161).chapter = 800
th_chose(162).TH_name = LoadResString_(3920, "")
th_chose(162).chose = 1
th_chose(162).text = LoadResString_(3925, "")
th_chose(162).chapter = 800
th_chose(163).TH_name = LoadResString_(3930, "")
th_chose(163).chose = 1
th_chose(163).text = th_chose(163).TH_name
th_chose(163).chapter = 800
th_chose(164).TH_name = LoadResString_(3935, "")
th_chose(164).chose = 1
th_chose(164).text = th_chose(164).TH_name
th_chose(164).chapter = 800
last_th_choose = 165
For i% = -6 To 180
regist_data.th_chose(i%) = th_chose(i%).chose
Next i%
''StatusBar1.Panels(1).Width = StatusBar1.Width - 1770
'监护密码异常
'For i% = 1 To 5
' ch = Mid$(protect_data.pass_word_for_teacher, i%, 1)
'  If ch < "0" Or ch > "z" Then
'   protect_data.pass_word_for_teacher = "00000"
''    MDIForm1.set_password.Checked = False
'     GoTo load_mark10
'  End If
'Next i%
'************************
'关闭信息库
'***********************
load_mark10:
display_information_string(1) = LoadResString_(1470, "") '直线上两点同名,请重新输入!"
display_information_string(2) = LoadResString_(1465, "") '"圆周上的点与圆心同名,请重新输入!!"
'Me.Timer1.Enabled = True
 Set C_wait_for_aid_point = New wait_for_aid_point
  Set C_display_wenti = New display_class
   Call C_display_wenti.Set_Me_object(C_display_wenti)
   Set C_display_wenti1 = New display_class
   Call C_display_wenti1.Set_Me_object(C_display_wenti1)
     Set C_curve = New curve_Class
     Set C_display_picture = New display_picture
     Call C_display_picture.set_me_class(C_display_picture)
  Call clear_wenti_display
   ' Call protect_code
     If event_statue <> exit_program Then
'*************************************************
     Call initial_set
      Call init_conditions(0)
       StrOpenFile = Command()
If StrOpenFile <> "" Then
  Call input_problem_from_file(StrOpenFile)
   Draw_form.Caption = Mid$(StrOpenFile, InStrRev(StrOpenFile, "\") + 1)
    Wenti_form.Caption = Draw_form.Caption
 Else
    If line_width = 0 Then
       line_width = 1
        condition_color = 3
         conclusion_color = 12
          fill_color = 7
      End If
path_and_file = ""
'初始化问题条件
last_char = 0
MDIForm1.Timer1.Enabled = False
Wenti_form.Picture1.Cls
Draw_form.Cls
event_statue = ready
'*********************************************************
End If
End If
End Sub

Private Sub MDIForm_Load()
Dim i%
Dim data_cond As String * 55
'标题<关于DDS平面几何>
App.Title = LoadResString_(2275, "\\1\\" + LoadResString_(110, ""))
Me.Caption = App.Title 'LoadResString_(110, "")
App.HelpFile = App.path & "\pmjh3.chm"
'*************
Timer2.Enabled = True
'测屏幕大小
screenx& = Screen.width - 80
screeny& = Screen.Height - 660
'设置主窗口大小
MDIForm1.width = screenx& ' Screen.Width
MDIForm1.Height = screeny& 'Screen.Height
'MDIForm1.Hide
'设置打印机
Text2.ZOrder 0
Text1.ZOrder 0
Draw_form.ZOrder 1
Wenti_form.ZOrder 1
Call Set_Mune_Item
load_error:
End Sub

Private Sub MDIForm_Resize()
'screenx& = MDIForm1.Width ' Screen.Width
'screeny& = MDIForm1.Height 'Screen.Height
'If MDIForm1.WindowState = 1 Then
'Unload exam_form '最小化时关闭MDIForm的非子窗口
'Unload Print_Form
'End If
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
Dim data  As String * 41
'On Error GoTo unload_error
If path_and_file <> "" And save_statue = 1 Then
If MsgBox(LoadResString_(1285, "\\1\\" + path_and_file), 4, "", "", 0) = 6 Then
 Call C_IO.save_prove_result(path_and_file)
End If
End If
Unload exam_form
Unload Print_Form
Unload inform
'Unload clinetdisplay
Close #1
Close #2
Call set_regist
End
unload_error:
End Sub

Private Sub mea_and_cal_Click()
mea_and_cal.visible = True
End Sub

Private Sub measure_Click()
     If Wenti_form.SSTab1.Tab <> 2 Then
     Wenti_form.SSTab1.Tab = 2
     Wenti_form.Picture3.Cls
      Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
      Call measur_again
     End If
End Sub

Private Sub method_Click()
If run_type > 0 And run_type < 11 Then
 Exit Sub
End If
method.Checked = True
method2.Checked = False
'method3.Checked = False
Wenti_form.VScroll1.visible = True
Wenti_form.Caption = wenti_form_title + LoadResString_(3955, "\\1\\" + LoadResString_(425, ""))
End Sub

Private Sub method0_Click()
Dim i%, temp_wenti_no%
Dim re As total_record_type
If run_type > 0 And run_type < 11 Then
 Exit Sub
End If
method1.Checked = False
method0.Checked = True
MDIForm1.re_name.Enabled = False
'MDIForm1.moldy_condition.Enabled = False
'MDIForm1.moldy_conclusion.Enabled = False
run_statue = 1 '12.10
'If prove_times = 0 Then
' prove_times = 1
'  Call init_prove0
'Else
' prove_times = 2
'  Call init_prove
'End If
prove_or_set_dbase = False
If wenti_type = 1 Then
If start_prove(0, 1, 0) = 1 Then
For i% = 0 To 3
 If conclusion_data(i%).no(0) > 0 Then
 Wenti_form.Picture1.CurrentY = 20 * (C_display_wenti.m_last_conclusion + 1)
  Wenti_form.Picture1.CurrentX = 0
  Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(15))
     Wenti_form.Picture1.Print LoadResString_(435, "") + "(   )"
  Wenti_form.Picture1.CurrentY = 20 * (C_display_wenti.m_last_conclusion + 1)
  Wenti_form.Picture1.CurrentX = 0
  Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(0))
     Wenti_form.Picture1.Print LoadResString_(435, "");
  Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(12))
  Wenti_form.Picture1.CurrentY = Wenti_form.Picture1.CurrentY + 2
    Wenti_form.Picture1.Print LoadResString_(3955, "\\1\\" + Chr(65 + i%));
  Wenti_form.Picture1.CurrentY = Wenti_form.Picture1.CurrentY - 2
  Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(0))
    ' Wenti_form.Picture1.Print ")"
Call set_display_string_no(conclusion_data(i%).ty, conclusion_data(i%).no(0), 0, 0)
Call arrange_display_no
  Wenti_form.Picture1.CurrentY = 20 * (C_display_wenti.m_last_conclusion + 6)
  Wenti_form.Picture1.CurrentX = 0
Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(0))
    Wenti_form.Picture1.Print LoadResString_(3960, "")
  Wenti_form.Picture1.CurrentY = 20 * (C_display_wenti.m_last_conclusion + 7)
  Wenti_form.Picture1.CurrentX = 0
Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(0))
Call set_display_string(True, 0, 0, 0, True)
'Call C_display_wenti.input_m_sentences(Wenti_form.Picture1, 1, _
           C_display_wenti.m_last_conclusion + 4, C_display_wenti.m_last_input_wenti_no, 1, False, 0)
End If
Next i%
End If
Else
If conclusion_data(0).ty > 0 Then
event_statue = wait_for_prove
set_or_prove = 0
display_type = 0
Call display_run(0)
ElseIf prove_times = 1 Then
Call MsgBox(LoadResString_(3965, ""), 64, "", 0, 0)
End If
End If
End Sub

Private Sub method1_Click()
Dim i%
If run_type > 0 And run_type < 11 Then
 Exit Sub
End If
method1.Checked = True
method0.Checked = False
MDIForm1.re_name.Enabled = False
run_statue = 3 '12.10
Call C_display_wenti.set_m_display_by_step(True)
prove_or_set_dbase = False
If conclusion_data(0).ty > 0 Then
set_or_prove = 0
display_type = 1
MDIForm1.Toolbar1.Buttons(19).visible = False
MDIForm1.Toolbar1.Buttons(17).visible = True
MDIForm1.Toolbar1.Buttons(18).visible = True
Call display_run(1)
Else
Call MsgBox(LoadResString_(3965, ""), 64, "", 0, 0)
End If

End Sub

Private Sub method2_Click()
Dim i%, j%
Dim ts$
Dim ch$
If run_type > 0 And run_type < 11 Then
 Exit Sub
End If
method.Checked = False
method2.Checked = True
'method3.Checked = False
MDIForm1.Toolbar1.Buttons(19).visible = True
MDIForm1.Toolbar1.Buttons(17).visible = False
MDIForm1.Toolbar1.Buttons(18).visible = False
Wenti_form.Caption = wenti_form_title & LoadResString_(3955, "\\1\\" + LoadResString_(3970, ""))
Wenti_form.VScroll1.visible = False
'Wenti_form.Label2.visible = 1
'Wenti_form.List1.visible = 1
'Wenti_form.List1.Clear
'Wenti_form.Label3.visible = 1
'Wenti_form.List2.visible = 1
'Wenti_form.List2.Clear
run_statue = 4 '12.10
'For i% = 0 To wenti_no - 1
'ts$ = ""
'For j% = 1 To Len(display_input_condition(i%).cond)
'ch$ = Mid$(display_input_condition(i%).cond, j%, 1)
'If ch$ <> "~" And ch$ <> "!" Then
'If ch$ = "#" Then
'ts$ = ts$ + "+"
'ElseIf ch$ = "#" Then
'ts$ = ts$ + "-"
'Else
'ts$ = ts$ + ch$
'End If
'End If
'Next j%
'Wenti_form.List1.AddItem ts$, i%
'Wenti_form.List1.Selected(i%) = True
'Next i%
'Load ch_ruler
'If ch_ruler.List2.ListCount > 0 Then
'For i% = 0 To ch_ruler.List2.ListCount - 1
'Wenti_form.List2.AddItem TH_CHOSE(ch_ruler.List2.ItemData(i%)).TH_name
'Wenti_form.List2.ItemData(i%) = ch_ruler.List2.ItemData(i%)
'Wenti_form.List2.Selected(i%) = True
'Next i%
'Else
'For i% = 1 To 160
'If TH_CHOSE(i%).chose = 1 Then
'Wenti_form.List2.AddItem TH_CHOSE(i%).TH_name
'Wenti_form.List2.ItemData(Wenti_form.List2.NewIndex) = i% 'ch_ruler.List2.ItemData(i%)
'Wenti_form.List2.Selected(Wenti_form.List2.NewIndex) = True
'End If
'Next i%
'End If
'List1.ItemData(List1.NewIndex) = List2.ItemData(List2.ListIndex)
'Wenti_form.List2 = ch_ruler.List2
End Sub

Private Sub method3_Click()
Dim i%
Exit Sub
method.Checked = False
method2.Checked = False
'method3.Checked = True
MDIForm1.re_name.Enabled = False
'MDIForm1.moldy_condition.Enabled = False
'MDIForm1.moldy_conclusion.Enabled = False
MDIForm1.conclusion.Enabled = True
MDIForm1.c_line1_5.Enabled = False
MDIForm1.c_cal.Enabled = False
MDIForm1.c_choose.Enabled = False
 Wenti_form.Caption = wenti_form_title & LoadResString_(3955, "\\1\\" + LoadResString_(3975, ""))
Wenti_form.VScroll1.visible = True
run_type = 4
run_type_1 = 3
draw_or_prove = 1
prove_or_set_dbase = False
If conclusion_data(0).ty > 0 Then
 event_statue = input_prove_by_hand
 ' For i% = 1 To last_dot_line
 'Call draw_dot_line(i%)
 'Next i%
'If wenti_no = c_display_wenti.m_last_conclusion + last_conclusion Then
'Wenti_form.Picture1.CurrentY = 20 * (c_display_wenti.m_last_conclusion + last_conclusion + 2)
'Wenti_form.Picture1.CurrentX = 0
'Wenti_form.Picture1.Print loadresstring_(450,"")
'End If
 set_or_prove = 1
' Call start_prove
 'input_sentence_no(0, 0) = 26
 'input_sentence_no(0, 1) = -30
 'input_sentence_no(0, 2) = -32
 'input_sentence_no(0, 3) = -33
 'input_sentence_no(1, 0) = -1
 'input_sentence_no(1, 1) = -6
 'input_sentence_no(1, 2) = -7
 'input_sentence_no(1, 3) = -34
 'input_sentence_no(1, 4) = -35
 'input_sentence_no(1, 5) = -40
 'input_sentence_no(2, 0) = 23
 'input_sentence_no(3, 0) = -36
 'input_sentence_no(3, 1) = -37
 'input_sentence_no(4, 0) = -5
 'input_sentence_no(4, 1) = -4
 'input_sentence_no(4, 2) = -38
 'input_sentence_no(4, 3) = -39
 'Else
'Call MsgBox(loadresstring_(712), 64, "", 0, 0)
End If


End Sub

Private Sub midpoint_Click()
Call menu_item_click("draw_point_and_line", 2, 2185)
End Sub


Private Sub moldy_conclusion_Click()
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(3980, "")
End Sub

Private Sub moldy_condition_Click()
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(3980, "")
End Sub

Private Sub modify_input_Click()
If modify_input_statue = False Then
 temp_problem_record = put_wenti_to_problem
  modify_input_statue = True
End If
End Sub

Private Sub move_part_Click()
operator = "change_picture"
Call remove_uncomplete_operat(old_operator)
Draw_form.Cls
Call draw_again1(Draw_form)
list_type_for_draw = 3
old_operator = operator
If last_conditions.last_cond(1).change_picture_type = line_ Then
 'Call draw_change_line(5)
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(3985, "")
ElseIf last_conditions.last_cond(1).change_picture_type = polygon_ Then
' Call draw_change_polygon(0)
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(3990, "")
ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
' Call draw_change_circle(0)
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(3995, "")
End If
'Draw_form.SetFocus
'protect_munu = 1
End Sub

Private Sub New_Click()
'If run_type > 0 And run_type < 11 Then
' Exit Sub
'End If
'Wenti_form.Show
If run_type > 10 Then
If path_and_file <> "" And save_statue = 1 Then
If MsgBox(LoadResString_(1285, "\\1\\" + path_and_file), 4, "", "", 0) = 6 Then
 Call C_IO.save_prove_result(path_and_file)
End If
ElseIf StrOpenFile <> "" And save_statue = 1 Then
    If MsgBox(LoadResString_(1285, "\\1\\" + Mid$(StrOpenFile, InStrRev(StrOpenFile, "\") + 1)), 4, "", "", 0) = 6 Then
 Call C_IO.save_prove_result(StrOpenFile)
End If
End If
End If
Call clear_wenti_display
Call init_conditions(0)
Call init_data_base
path_and_file = ""
'初始化问题条件
last_char = 0
MDIForm1.Timer1.Enabled = False
Wenti_form.Picture1.Cls
Draw_form.Cls
event_statue = ready
End Sub


Private Sub Open_Click()
Dim mark As String * 8
Dim i%, m%, j%, k%
Dim temp_ch As String * 1
Dim temp_string As String * 16
Dim temp_string0 As String * 16 'path_and_file As String
Dim temp_result As String
Dim of_struct As OFSTRUCT

If run_type > 0 And run_type < 11 Then
 Exit Sub
End If
If path_and_file <> "" And save_statue = 1 Then
If MsgBox(LoadResString_(1285, "\\1\\" + path_and_file), 4, "", "", 0) = 6 Then
 Call C_IO.save_prove_result(path_and_file)
End If
End If
Call init_conditions(0)
'CommonDialog1.CancelError = True
'On Error GoTo open_click_error
'CommonDialog1.filer = ""
CommonDialog1.ShowOpen
path_and_file = CommonDialog1.FileName
'if commondialog1.
If path_and_file <> "" Then
 Call input_problem_from_file(path_and_file)
End If
Exit Sub
open_click_error:
End Sub

Private Sub pandline_Click() '子菜单 画点和线
operator = "draw_point_and_line"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 1
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2180, "")
'Draw_form.SetFocus
Me.ZOrder
End Sub

Private Sub paral_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & paral.Caption)
inform.Picture1.Cls
Call set_information_list(paral_)
End Sub


Private Sub paralandverti_Click()
  Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 1
   draw_step = 0
    Call init_draw_data
     old_operator = operator
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4000, "")
  Draw_form.HScroll1.visible = False
End Sub
Private Sub mprint_Click()
Dim i%
'Load P_Form
If C_display_wenti.m_last_input_wenti_no > 0 Then
If protect_data.pass_word_for_teacher = "00000" Or _
    InStr(1, last_conditions.last_cond(1).pass_word_for_teacher, "*", 0) = 0 Then
If run_type > 0 And run_type < 11 Then
 Exit Sub
End If
Print_Form.Show
Print_Form.print1.Enabled = print_enabled
Print_Form.Picture1.SetFocus
Printer.KillDoc
p_x% = 2000
p_y% = 2000
text_y% = 5500
page_note_string = LoadResString_(110, "")
For i% = 1 To last_conditions.last_cond(1).point_no
 point_name_position(i%) = 0
Next i%
'page_no% = (wenti_no - 21) / 37 + 1
Call C_display_wenti.m_Print_wenti(Print_Form.Picture1, 0, 1, 0, True)
End If
End If
End Sub

Private Sub porine_Click()
'主菜单 点和线
End Sub

Private Sub ratio_point_Click()
Call menu_item_click("draw_point_and_line", 3, 2190)
End Sub
Private Sub menu_item_click(opera As String, draw_type As Byte, dis_no As Integer)
'protect_munu = 1
'protect_munu_ = 1
operator = opera '"draw_point_and_line"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = draw_type
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(dis_no, "")
'If protect_munu_ = 1 Then
' protect_munu = 0
'End If
Call control_menu(False) '关闭一些菜单
End Sub
Private Sub re_line_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & re_line)
inform.Picture1.Cls
Call set_information_list(dpoint_pair_)

End Sub

Private Sub re_name_all_Click()
Dim i%
operator = "re_name"
list_type_for_draw = 1
'Call C_display_picture.m_draw_point
  MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4005, "")
'Call C_display_picture.m_draw_point
 'Call C_display_wenti.input_m_sentences(Wenti_form.Picture1, 0, 1, _
       C_display_wenti.m_last_input_wenti_no, 0, False, 0)
 Call init_input0
 choose_point = 1
  For i% = 1 To last_conditions.last_cond(1).point_no
   Call set_point_name(i%, "")  'loadresstring_(1226,"")
 Next i%
 'Call C_display_picture.draw_red_point(choose_point)
 Call C_display_picture.flash_point(choose_point)
 'Call C_display_wenti.input_m_sentences(Wenti_form.Picture1, 1, 1, _
       C_display_wenti.m_last_input_wenti_no, 0, False, 0)
 re_name_ty = 1
End Sub

Private Sub re_name_one_Click()
Dim i%
operator = "re_name"
list_type_for_draw = 2
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4010, "")
For i% = 1 To last_conditions.last_cond(1).point_no
 If m_poi(i%).data(0).data0.visible = 0 Then
  Call set_point_name(i%, "")
 End If
Next i%
'Draw_form.SetFocus
'protect_munu = 1
End Sub


Private Sub relation_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & relation.Caption)
inform.Picture1.Cls
Call set_information_list(relation_)
End Sub


Private Sub remove_point__Click()
operator = "re_name"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 3
chose1.Caption = LoadResString_(4015, "")
chose2.Caption = LoadResString_(4020, "")
chose3.visible = False
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4025, "")
'Draw_form.SetFocus
Me.ZOrder
Dim w_n%
    Call C_display_wenti.set_m_conclusion_or_condition(w_n%, condition)
    Call C_display_wenti.set_m_string("", "yeyiq", "", "", "(urwro)", 0, -1, w_n%, 0)
End Sub

Private Sub Right_angle_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & right_angle.Caption)
inform.Picture1.Cls
Call set_information_list(Rangle_)
End Sub
Private Sub Run_Click(index As Integer)
If run_type = 0 Or run_type > 10 Then
Call clear_wenti_display
Call remove_uncomplete_operat(old_operator)
'Call init_operat
End If
End Sub

Private Sub run_type0_Click()
If run_type > 0 And run_type < 11 Then
 Exit Sub
End If
'run_type0.Checked = True
'run_type1.Checked = False
'run_type2.Checked = False
run_type = 0
run_type_1 = 0
Toolbar1.Buttons(19).visible = True
method.Enabled = True
'MDIForm1.method.checked = True
method2.Enabled = True
'method3.Enabled = True
End Sub

Private Sub run_type1_Click()
If run_type > 0 And run_type < 11 Then
 Exit Sub
End If
'run_type1.Checked = True
'run_type0.Checked = False
'run_type2.Checked = False
run_type = 1
run_type_1 = 1
Toolbar1.Buttons(19).visible = True
method.Enabled = True
method2.Enabled = True
'method3.Enabled = True
'add_point.Enabled = True
End Sub

Private Sub run_type2_Click()
If run_type > 0 And run_type < 11 Then
 Exit Sub
End If
'run_type2.Checked = True
'run_type0.Checked = False
'run_type1.Checked = False
run_type = 1
run_type_1 = 2
Toolbar1.Buttons(19).visible = True
method.Enabled = True
method2.Enabled = True
'method3.Enabled = True
End Sub
Private Sub S11_Click()
Call input_sentence_y(1, 0, 0)
End Sub
Private Sub S12_Click()
Call input_sentence_y(1, -1, 1)
End Sub
Private Sub S13_Click()
Call input_sentence_y(1, -24, 0)
End Sub
Private Sub S14_Click()
Call input_sentence_y(1, -4, 1)
End Sub
Private Sub S15_Click()
Call input_sentence_y(1, -5, 1)
End Sub
Private Sub S16_Click()
Call input_sentence_y(1, -6, 1)
End Sub
Private Sub S17_Click()
Call input_sentence_y(1, -7, 1)
End Sub
Private Sub S18_Click()
Call input_sentence_y(1, -34, 1)
End Sub
Private Sub S19_Click()
Call input_sentence_y(1, -40, 1)
End Sub
Private Sub S1A_Click()
Call input_sentence_y(1, -35, 1)
End Sub
Private Sub S1B_Click()
Call input_sentence_y(1, -41, 1)
End Sub
Private Sub s1C_Click()
wenti_type = 0
Call input_sentence_y(1, -49, 1)
End Sub
Private Sub s1D_Click()
Call input_sentence_y(1, -50, 0)
End Sub
Private Sub s1E_Click()
Call input_sentence_y(1, -51, 0)
End Sub
Private Sub s1F_Click()
Call input_sentence_y(1, -52, 0)
End Sub
Private Sub s1G_Click()
Call input_sentence_y(1, 22, 1)
End Sub
Private Sub S21_Click()
Call input_sentence_y(1, 1, 0)
End Sub
Private Sub S22_Click()
Call input_sentence_y(1, 2, 0)
End Sub
Private Sub S23_Click()
Call input_sentence_y(1, 3, 0)
End Sub
Private Sub S24_Click()
Call input_sentence_y(1, 4, 0)
End Sub
Private Sub S31_Click()
Call input_sentence_y(1, 5, 0)
End Sub
Private Sub S32_Click()
Call input_sentence_y(1, 6, 0)
End Sub
Private Sub S33_Click()
If event_statue = wait_for_input_char Or event_statue = _
    wait_for_modify_char Or event_statue = input_char_again Then
     Wenti_form.Picture1.SetFocus
      Exit Sub
ElseIf event_statue = wait_for_draw_point Then
     'Draw_form.SetFocus
    ' protect_munu = 1
      Exit Sub
Else
'inp = 9
Call input_sentence_y(1, 9, 0)
End If
End Sub
Private Sub S34_Click()
Call input_sentence_y(1, 14, 0)
End Sub
Private Sub S35_Click()
Call input_sentence_y(1, -22, 0)
End Sub
Private Sub S36_Click()
Call input_sentence_y(1, -23, 0)
End Sub
Private Sub S37_Click()
Call input_sentence_y(1, 10, 0)
End Sub
Private Sub S38_Click()
Call input_sentence_y(1, 16, 0)
End Sub
Private Sub S39_Click()
Call input_sentence_y(1, -31, 0)
End Sub
Private Sub S3A_Click()
Call input_sentence_y(1, -43, 0)
End Sub
Private Sub S41_Click()
Call input_sentence_y(1, 7, 0)
End Sub
Private Sub S42_Click()
Call input_sentence_y(1, -71, 0)
End Sub
Private Sub S43_Click()
Call input_sentence_y(1, 11, 0)
End Sub
Private Sub S44_Click()
Call input_sentence_y(1, 13, 0)
End Sub
Private Sub S45_Click()
Call input_sentence_y(1, -33, 0)
End Sub
Private Sub S46_Click()
Call input_sentence_y(1, -32, 0)
End Sub
Private Sub S47_Click()
Call input_sentence_y(1, -3, 0)
End Sub
Private Sub S48_Click()
Call input_sentence_y(1, -2, 0)
End Sub
Private Sub S49_Click()
Call input_sentence_y(1, -30, 0)
End Sub
Private Sub S4A_Click()
Call input_sentence_y(1, -24, 0)
End Sub
Private Sub S4B_Click()
Call input_sentence_y(1, -42, 0)
End Sub
Private Sub S4C_Click()
Call input_sentence_y(1, 12, 0)
End Sub
Private Sub S51_Click()
Call input_sentence_y(1, 18, 1)
End Sub
Private Sub S52_Click()
Call input_sentence_y(1, 19, 1)
End Sub
Private Sub S53_Click()
Call input_sentence_y(1, 20, 1)
End Sub
Private Sub S54_Click()
Call input_sentence_y(1, 21, 1)
End Sub
Private Sub S61_Click()
Call input_sentence_y(1, -20, 1)
End Sub
Private Sub S62_Click()
Call input_sentence_y(1, -19, 1)
End Sub
Private Sub S63_Click()
Call input_sentence_y(1, -18, 1)
End Sub
Private Sub S64_Click()
Call input_sentence_y(1, -17, 1)
End Sub
Private Sub S65_Click()
Call input_sentence_y(1, -16, 1)
End Sub
Private Sub S66_Click()
Call input_sentence_y(1, -15, 1)
End Sub
Private Sub S67_Click()
Call input_sentence_y(1, -14, 1)
End Sub
Private Sub S68_Click()
Call input_sentence_y(1, -13, 1)
End Sub
Private Sub S69_Click()
Call input_sentence_y(1, -12, 1)
End Sub
Private Sub S6A_Click()
Call input_sentence_y(1, -11, 1)
End Sub
Private Sub S6B_Click()
Call input_sentence_y(1, -10, 1)
End Sub
Private Sub S6C_Click()
Call input_sentence_y(1, -9, 1)
End Sub

Private Sub S6D_Click()
Call input_sentence_y(1, -8, 1)
End Sub
Private Sub S6E_Click()
Call input_sentence_y(1, -48, 1)
End Sub
Private Sub S6F_Click()
Call input_sentence_y(1, -47, 1)
End Sub
Private Sub S6G_Click()
Call input_sentence_y(1, -46, 1)
End Sub
Private Sub S6H_Click()
Call input_sentence_y(1, -45, 1)
End Sub
Private Sub S6I_Click()
Call input_sentence_y(1, -37, 1)
End Sub
Private Sub S6J_Click()
Call input_sentence_y(1, -36, 1)
End Sub
Private Sub save_as_Click()
Dim mark As String * 8
Dim i%
Dim ts$
Dim temp_path_and_file As String
  'On Error GoTo save_as_click_error
If run_type > 0 And run_type < 11 Then
 Exit Sub
End If
CommonDialog1.ShowSave
   temp_path_and_file = CommonDialog1.FileName
If path_and_file <> "" Then
    ts$ = path_and_file + LoadResString_(4030, "")
     If DoesFileExist(temp_path_and_file) Then
      If MsgBox(ts$, vbYesNo, "", 0, 0) = 6 Then
       path_and_file = CommonDialog1.FileName
         Call CopyFile(temp_path_and_file, path_and_file, 1)
          path_and_file = temp_path_and_file
      End If
     Else
       If CopyFile(temp_path_and_file, path_and_file, 0) = 0 Then
          Call CopyFile(temp_path_and_file, path_and_file, 1)
          path_and_file = temp_path_and_file
       End If
     End If
End If
save_as_click_error:
End Sub

Private Sub save_Click()
Dim mark As String * 8
Dim i%
Dim ts$
If run_type > 0 And run_type < 11 Then
 Exit Sub
End If
If path_and_file = "" Then
' CommonDialog1.CancelError = True
  'On Error GoTo save_click_error
 CommonDialog1.ShowSave
   path_and_file = CommonDialog1.FileName
   If path_and_file <> "" Then
   ts$ = path_and_file + LoadResString_(4030, "")
   If DoesFileExist(path_and_file) Then '文件存在
      ts$ = path_and_file + LoadResString_(4030, "")
      If MsgBox(ts$, vbYesNo, "", 0, 0) = 6 Then
        Call C_IO.save_prove_result(path_and_file)
      End If
     Else
        Call C_IO.save_prove_result(path_and_file)
     End If
   End If
Else
  'If path_and_file <> "" Then
   ts$ = path_and_file + LoadResString_(4030, "")
    If DoesFileExist(path_and_file) Then
      ts$ = path_and_file + LoadResString_(4030, "")
      If MsgBox(ts$, vbYesNo, "", 0, 0) = 6 Then
        Call C_IO.save_prove_result(path_and_file)
      End If
     Else
        Call C_IO.save_prove_result(path_and_file)
     End If
   'End If

 'CommonDialog1.CancelError = True
 'CommonDialog1.ShowSave
 'path_and_file = CommonDialog1.FileName
 'If path_and_file <> "" Then
 'ts$ = path_and_file + loadresstring_(728)
 'If DoesFileExist(path_and_file) Then
 ' If MsgBox(ts$, vbYesNo, "", 0, 0) Then
 'Call save_prove_result(path_and_file)
 ' End If
 'Else
 'Call save_prove_result(path_and_file)
 'End If
 'End If
 End If
save_click_error:
 End Sub

Private Sub savee_Click()
If run_type > 0 And run_type < 11 Then
 Exit Sub
Else
  Load IO_form
  IO_form.Show
  io_statue = 1
End If
End Sub

Private Sub set_area__of_polygon_Click()
operator = "set"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 4
old_operator = operator
Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4035, "")
'Draw_form.SetFocus

End Sub

Private Sub set_change_line_Click()
operator = "change_picture"
Call remove_uncomplete_operat(old_operator)
'***Call set_old_picture
list_type_for_draw = 6
old_operator = operator
line_for_change.similar_ratio = 1
line_for_change.move.X = 0
line_for_change.move.Y = 0
line_for_change.line_no(0).center(0).X = 0
line_for_change.line_no(0).center(0).Y = 0
line_for_change.line_no(0).center(1) = line_for_change.line_no(0).center(0)
line_for_change.line_no(0).in_point(0) = 0
line_for_change.line_no(0).coord(0).X = 0
line_for_change.line_no(0).coord(0).Y = 0
line_for_change.line_no(0).coord(1) = line_for_change.line_no(0).coord(0)
line_for_change.rote_angle = 0
 Call BitBlt(Draw_form.Picture1.hdc, 0, 0, Draw_form.Picture1.width, _
     Draw_form.Picture1.Height, Draw_form.hdc, 0, 0, vbSrcCopy)
       picture_copy = True '原图Draw_form.hdc复制到Draw_form.Picture1.hdc
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1955, "")
'Draw_form.SetFocus
'protect_munu = 1
Me.ZOrder
End Sub

Private Sub set_change_type_Click()
If set_change_type_ = False Then
If line_width = 1 Then
line_width = 2
Draw_form.DrawWidth = 2
End If
If last_conditions.last_cond(1).change_picture_type = polygon_ Then
'***Call draw_change_polygon(0)
ElseIf last_conditions.last_cond(1).change_picture_type = circle_ Then
'***Call draw_change_circle(0)
ElseIf last_conditions.last_cond(1).change_picture_type = line_ Then
End If
set_change_type_ = True
End If
End Sub

Private Sub set_circle_for_change_Click()
operator = "change_picture"
Call remove_uncomplete_operat(old_operator)
'***Call set_old_picture
list_type_for_draw = 2
old_operator = operator
Circle_for_change.similar_ratio = 1
Circle_for_change.c = 0
Circle_for_change.c_coord.X = 0
Circle_for_change.c_coord.Y = 0
Circle_for_change.move.X = 0
Circle_for_change.move.X = 0
Circle_for_change.rote_angle = 0
Circle_for_change.radii = 0
Circle_for_change.direction = 1
   Call BitBlt(Draw_form.Picture1.hdc, 0, 0, Draw_form.Picture1.width, _
     Draw_form.Picture1.Height, Draw_form.hdc, 0, 0, vbSrcCopy)
       picture_copy = True '原图Draw_form.hdc复制到Draw_form.Picture1.hdc
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1110, "")
'Draw_form.SetFocus
'protect_munu = 1
End Sub

Private Sub set_dis_p_line_Click()
'Call window_fore_rear
operator = "set"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 3
old_operator = operator
Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4040, "")
'Draw_form.SetFocus

End Sub

Private Sub set_for_measure_Click()
     If Wenti_form.SSTab1.Tab <> 2 Then
     Wenti_form.SSTab1.Tab = 2
     Wenti_form.Picture3.Cls
      Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
      Call measur_again
     End If
End Sub

Private Sub set_length_Click()
'Call window_fore_rear
operator = "set"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 2
old_operator = operator
Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4045, "")
'Draw_form.SetFocus
'protect_munu = 1
End Sub

Private Sub set_mode_Click()
set_form.Show
Call initial_set
End Sub

Private Sub set_opera_type_Click()
'If operator = "move_point" And _
       last_conditions.last_cond(1).last_view_point_no > 0 Then
'        operator = "move_point_"
'ElseIf last_conditions.last_cond(1).last_view_point_no = 0 Then
 operator = "move_point"
'  Draw_form.Picture1.visible = False
   Call remove_uncomplete_operat(old_operator)
'End If
End Sub

Private Sub set_p_Click()
widthform.Show
temp_line_width = line_width
temp_condition_color = condition_color
temp_conclusion_color = conclusion_color
temp_fill_color = fill_color
Call init_set
End Sub

Private Sub set_polygon_for_change_Click()
operator = "change_picture"
Call remove_uncomplete_operat(old_operator)
'***Call set_old_picture
list_type_for_draw = 1
old_operator = operator
Polygon_for_change.move.X = 0
Polygon_for_change.move.Y = 0
Polygon_for_change.direction = 1
Polygon_for_change.rote_angle = 0
Polygon_for_change.similar_ratio = 1
Polygon_for_change.p(0).direction = True
Polygon_for_change.p(0).center.X = 0
Polygon_for_change.p(0).center.Y = 0
Polygon_for_change.p(0).coord_center.X = 0
Polygon_for_change.p(0).coord_center.Y = 0
Polygon_for_change.p(0).direction = True
Polygon_for_change.p(0).total_v = 0
Polygon_for_change.p(1) = Polygon_for_change.p(0)
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2320, "")
Polygon_for_change.similar_ratio = 1
 Call BitBlt(Draw_form.Picture1.hdc, 0, 0, Draw_form.Picture1.width, _
     Draw_form.Picture1.Height, Draw_form.hdc, 0, 0, vbSrcCopy)
       picture_copy = True '原图Draw_form.hdc复制到Draw_form.Picture1.hdc
'Draw_form.SetFocus
'protect_munu = 1
End Sub

Private Sub set_ruler_Click()
'Call window_fore_rear
operator = "set"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 1
old_operator = operator
      If Ratio_for_measure.Ratio_for_measure = 0 Then '第一次设置测量标尺
         Ratio_for_measure.Ratio_for_measure = 20
      End If
Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
Wenti_form.HScroll2.visible = True
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1839, "")
'Draw_form.SetFocus
End Sub

Private Sub set_view_point_Click()
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 0
old_operator = operator
operator = "set_view_point"
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4050, "")
'Draw_form.Picture1.visible = True
'Draw_form.Picture1.Top = 0
'Draw_form.Picture1.Left = 0
'Call BitBlt(Draw_form.Picture1.hdc, 0, 0, Draw_form.Picture1.Width, _
        Draw_form.Picture1.Height, Draw_form.hdc, 0, 0, vbSrcCopy) '将原图存入Picture1
'Call SetWindowLong(Draw_form.Picture1.hwnd, -20, &H20&)
' Draw_form.Picture1.Refresh
' Draw_form.SetFocus
End Sub
Private Sub similar_triangle_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & similar_triangle.Caption)
inform.Picture1.Cls
Call set_information_list(similar_triangle_)
End Sub
Private Sub solve_Click()
If run_type > 0 And run_type < 11 Then
 Exit Sub
End If
Call run_
End Sub
Private Sub sp_four_sides_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & sp_four_sides.Caption)
inform.Picture1.Cls
Call set_information_list(sp_polygon4_)
End Sub
Private Sub StatusBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ts$
 'If Button = 1 Then
   If Button = 2 Then
 If X > 0 And X < 500 Then
 Me.StatusBar1.Align = 1 + Me.StatusBar1.Align Mod 2
 'If X > 50 And X < 5000 Then
 ' If test_no% = 0 Then
 '  test_no% = 10
 ' ElseIf test_no% = 10 Then
 '  test_no% = 20
 ' ElseIf test_no% = 20 Then
 '  test_no% = 30
 ' Else
 '  test_no% = 5
 ' End If
 ElseIf X > 5000 And X < 5200 Then
 ts$ = LoadResString_(4055, "") + protect_data.pass_word + LoadResString_(1240, "") + _
        protect_data.pass_word_for_teacher _
          + "; software_regist_no:2001SR3847" 'cal_password(protect_data.serial_no, protect_data.computer_id)
 Call MsgBox(ts$, vbDefaultButton2, "", 0, 0)
 ElseIf X > 9500 And X < 9700 Then
 Call MsgBox(LoadResString_(2260, ""), vbDefaultButton2, LoadResString_(4060, ""), 0, 0)
 ElseIf X < 50 Then
 Else
  test_no% = 0
 End If
End If
End Sub

Private Sub sum_two_angle_pi_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & sum_two_angle_pi.Caption)
inform.Picture1.Cls
Call set_information_list(two_angle_180)
End Sub

Private Sub sum_two_angle_right_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & sum_two_angle_right.Caption)
inform.Picture1.Cls
Call set_information_list(angle2_right)
End Sub


Private Sub tangent_circle_Click()
operator = "draw_circle"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 5
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1665, "")
'Draw_form.SetFocus
'protect_munu = 1
End Sub

Private Sub tangent_line_Click()
operator = "draw_circle"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 3
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4065, "")
'Draw_form.SetFocus
'protect_munu = 1
End Sub

Private Sub tangent_line_point_to_circle_Click()
operator = "draw_circle"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 3
old_operator = operator
'MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1665, "")
'Draw_form.SetFocus
'protect_munu = 1

End Sub

Private Sub tangent_of_two_circles_Click()
operator = "draw_circle"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 4
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(1665, "")
'Draw_form.SetFocus
'protect_munu = 1

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim s!
Dim k% 'text_num!
Dim l% ', text_num1%, text_num2%
Dim ch$
Dim p_c As POINTAPI
Dim t_c As circle_data_type
If Text1.text = "" Then
 If KeyAscii = 13 Then
  reduce_level0 = val(Text2.text)
   Text1.visible = False
    Text2.visible = False
 Else
  Text2.text = Text2.text & Chr(KeyAscii)
 End If
End If
MDIForm1.StatusBar1.Panels(1).text = ""
If KeyAscii = 13 Then
input_text_statue = False
 input_text_finish = True
    If InStr(1, Text2.text, ".", 0) > 0 Then
     text_num! = val(Text2.text)
    Else
     k% = InStr(1, Text2.text, "/", 0)
     If k% = 0 Then
      k% = InStr(1, Text2.text, ":", 0)
       If k% = 0 Then
             text_num1% = val(Text2.text)
              text_num2% = 1
               text_num! = text_num1%
       End If
    End If
  If k% > 0 Then
   text_num1% = val(Mid$(Text2.text, 1, k% - 1))
    text_num2% = val(Mid$(Text2.text, k% + 1, Len(Text2.text)))
     text_num! = text_num1% / text_num2%
   End If
  Call simple_two_int(text_num1%, text_num2%)
 End If
Select Case operator
Case "draw_point_and_line"
If list_type_for_draw = 3 Then
ElseIf list_type_for_draw = 4 Then
 text_num! = val(Text2.text)
    temp_circle(0) = m_circle_number( _
      1, temp_point(0).no, pointapi0, 0, 0, 0, _
        CLng(Ratio_for_measure.Ratio_for_measure * text_num!), 0, 0, 1, 1, condition, fill_color, True)
         m_Circ(temp_circle(0)).data(0).parent.element(0).no = m_Circ(temp_circle(0)).data(0).data0.center
 '         Call set_circle_in_point(temp_circle(0), 3, CInt(text_num! * 100), condition)
 '          Call set_circle_in_point(temp_circle(0), 2, -1, condition)
 MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4070, "")
      draw_step = 2
      End If
Case "set"
 If list_type_for_draw = 2 Then
       Call Wenti_form.Picture3.Cls
        Ratio_for_measure.Ratio_for_measure = length_(set_measure_no%).len / text_num!
         Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
          Call change_ratio
           Call measur_again
     Text1.visible = False
      Text2.visible = False
  ElseIf list_type_for_draw = 3 Then
      Call Wenti_form.Picture3.Cls 'draw_ruler(ratio_for_measure.ratio_for_measure, delete)
    Ratio_for_measure.Ratio_for_measure = Abs(length_point_to_line(set_measure_no%).len) / text_num!
      Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
        Call change_ratio
         Call measur_again
     Text1.visible = False
      Text2.visible = False
  ElseIf list_type_for_draw = 4 Then
      Call Wenti_form.Picture3.Cls 'draw_ruler(ratio_for_measure.ratio_for_measure, delete)
    Ratio_for_measure.Ratio_for_measure = sqr(Abs(Area_polygon(set_measure_no%).Area) / text_num!)
      Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
        Call change_ratio
         Call measur_again
     Text1.visible = False
      Text2.visible = False
 End If
End Select
MDIForm1.Text1.visible = False
 MDIForm1.Text2.visible = False

End If

End Sub


Private Sub three_angle_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & three_angle.Caption)
inform.Picture1.Cls
Call set_information_list(angle3_value_)
End Sub


Private Sub Three_point_on_line_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & three_point_on_line.Caption)
inform.Picture1.Cls
Call set_information_list(point3_on_line_)
End Sub

Private Sub Timer1_Timer()
If event_statue = wait_for_draw_point Then
 event_statue = wait_for_input_char
End If
If event_statue = -1 Then
Call init_conditions(0)
Timer1.interval = 400
event_statue = 0
Timer1.Enabled = False
ElseIf time11_display_type = icon_ Then 'C_display_wenti.m_icon_display Then
   If time_no = 0 Then
    Call C_display_wenti.display_icon(0)
   Else
    Call C_display_wenti.display_icon(1)
   End If
    time_no = (time_no + 1) Mod 2  '明暗交替
ElseIf time11_display_type = inform_ <> 0 Then
   If time_no = 0 Then
      Call draw_picture_for_inform_(0)
   Else
      Call draw_picture_for_inform_(1)
   End If
    time_no = (time_no + 1) Mod 2  '明暗交替
ElseIf event_statue = wait_for_modify_sentence Then
  If time_no = 0 Then
   'Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, 1, modify_wenti_no, _
            display_wenti_h_position%, 1, 0, 0)
    Else
     'Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, 0, modify_wenti_no, _
            display_wenti_h_position%, 1, 0, 0)
  End If
   time_no = (time_no + 1) Mod 2
ElseIf event_statue = wait_for_prove Then
  If time_no = 0 Then
    'Draw_form.label1.ForeColor = QBColor(12) ' = "正在进行推理，请等待！"
    MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4075, "")
  Else
    'Draw_form.label1.ForeColor = QBColor(15) 'Caption = "正在进行推理，请等待！"
    MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4075, "")
 End If
  time_no = (time_no + 1) Mod 2
ElseIf event_statue = ready Then
 MDIForm1.StatusBar1.Panels(1).text = ""
Else
   If time_no = 0 Then
    Call C_display_wenti.display_icon(0)
   Else
    Call C_display_wenti.display_icon(1)
   End If
    time_no = (time_no + 1) Mod 2  '明暗交替
End If
time_act = True '计时器发出时信号
End Sub
Private Sub Timer2_Timer()
pro_no% = pro_no% + 1
data_no% = data_no% + 1
If run_type = 0 And event_statue = wait_for_prove Then
 pro_no1% = pro_no1% + 1
End If
If run_type = 0 And pro_no1% = 2 Then '30 Then
 run_type = 1
  Call find_conclusion1(0, 0, False)
'elseif event_statue=exit_program and
End If
'*********
If (data_no% >= 1 And protect_data.install_statue <> "S" _
                    And protect_data.install_statue <> "T") Or data_no% >= 12 Then
'一分钟计时
Call calculate_run_time
data_no% = 0
'If event_statue = exit_program Then
'  Call MsgBox("欢迎您使用《DSH－平面几何》测试版。测试版可以累计使用300分钟,其后每次限时使用６0秒。《DSH－平面几何》由中国科学院成都计算机应用研究所丁孙荭研究员研究开发。 ", 64, "申  明", 0, 0)
' End
'End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As Button)
Dim i%, temp_wenti_no%
Dim path As String
'Draw_form.Picture1.visible = False
 '防止点击菜单时,产生画图动作
If picture_copy = True Then
 Draw_form.Picture1.visible = False
 Call Draw_form.Cls
 Call BitBlt(Draw_form.hdc, 0, 0, Draw_form.Picture1.width, _
     Draw_form.Picture1.Height, Draw_form.Picture1.hdc, 0, 0, vbSrcCopy)
picture_copy = False
End If
Select Case Button.Key
Case "new"
 If run_type > 0 And run_type < 11 Then
  Exit Sub
 End If
 Call New_Click
Case "open"
 If run_type > 0 And run_type < 11 Then
  Exit Sub
 End If
 Call Open_Click
 Case "save"
 If run_type > 0 And run_type < 11 Then
  Exit Sub
 End If
 Call save_Click
Case "print"
 If run_type > 0 And run_type < 11 Then
  Exit Sub
 End If
 Call mprint_Click
Case "point"
 If run_type > 0 And run_type < 11 Then
  Exit Sub
 End If
 operator = "draw_point_and_line"
 If event_statue <> complete_input Then
 PopupMenu porine, 0, Button.left, Button.top + 350
 End If
Case "circle"
 If run_type > 0 And run_type < 11 Then
  Exit Sub
 End If
 operator = "draw_circle"
 PopupMenu draw_circle, 0, Button.left, Button.top + 350
Case "paral"
 If run_type > 0 And run_type < 11 Then
  Exit Sub
 End If
  operator = "paral_verti"
  PopupMenu paralandverti0, 0, Button.left, Button.top + 350
Case "poly"
 If run_type > 0 And run_type < 11 Then
  Exit Sub
 End If
 operator = "epolygon"
 PopupMenu E_polygon, 0, Button.left, Button.top + 350
Case "anima"
 If run_type > 0 And run_type < 11 Then
  Exit Sub
 End If
 ' If operator <> "move_point" And _
 '     last_conditions.last_cond(1).last_view_point_no > 0 Then
 '   operator = "move_point_"
 ' Else
    operator = "move_point"
 ' End If
  PopupMenu move_picture, 0, Button.left, Button.top + 350
Case "change"
 If run_type > 0 And run_type < 11 Then
  Exit Sub
 End If
  PopupMenu change_picture, 0, Button.left, Button.top + 350
   If operator = "change_picture" Then
    Else
   operator = "change_picture"
   End If
Case "measur"
 If run_type > 0 And run_type < 11 Then
  Exit Sub
 End If
  operator = "measure"
    PopupMenu measure, 0, Button.left, Button.top + 350
     If Wenti_form.SSTab1.Tab <> 2 Then
     Wenti_form.SSTab1.Tab = 2
     Wenti_form.Picture3.Cls
      Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
      Call measur_again
     End If
Case "set"
 If run_type > 0 And run_type < 11 Then
  Exit Sub
 End If
  operator = "set"
    PopupMenu set_for_measure, 0, Button.left, Button.top + 350
     If Wenti_form.SSTab1.Tab <> 2 Then
     Wenti_form.SSTab1.Tab = 2
     Wenti_form.Picture3.Cls
      If Ratio_for_measure.Ratio_for_measure = 0 Then '第一次设置测量标尺
         Ratio_for_measure.Ratio_for_measure = 20
      End If
      Call draw_ruler(Ratio_for_measure.Ratio_for_measure, display)
      Call measur_again
     End If
Case "ask"
 If run_type > 0 And run_type < 11 Then
  Exit Sub
 End If
  operator = "ask" 'LoadResString_(140)
  Call remove_uncomplete_operat(old_operator)
     old_operator = operator
   MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4080, "")
Case "name"
 If run_type > 0 And run_type < 11 Then
  Exit Sub
 End If
  PopupMenu re_name, 0, Button.left, Button.top + 350
  operator = "re_name"
  If C_display_wenti.m_last_input_wenti_no > C_display_wenti.m_last_conclusion And _
                                           C_display_wenti.m_last_conclusion > 0 Then
  Exit Sub
  End If
Case "postpone"
If contro_process <> postpone_ Then
    contro_process = postpone_
  StatusBar1.Panels(1).text = LoadResString_(4095, "") + _
                          "." + LoadResString_(4100, "\\1\\" + str(last_conditions.last_cond(1).total_condition))
   MDIForm1.Toolbar1.Buttons(17).Image = 18 'i% - 2
Else
 event_statue = proving
  If picture_copy = True Then
   Draw_form.Cls
    Call BitBlt(Draw_form.hdc, 0, 0, Draw_form.Picture1.width, _
     Draw_form.Picture1.Height, Draw_form.Picture1.hdc, 0, 0, vbSrcAnd)
      picture_copy = False
  End If
      MDIForm1.Toolbar1.Buttons(17).Image = 15 'i% - 2
  StatusBar1.Panels(1).text = LoadResString_(4100, "\\1\\" + str(last_conditions.last_cond(1).total_condition))
    contro_process = 0
End If
Case "stop"
 contro_process = stop_
  StatusBar1.Panels(1).text = LoadResString_(4105, "")
   run_type = 11
Case "run"
  MDIForm1.examp.Enabled = False
   contro_process = 0
   MDIForm1.Toolbar1.Buttons(17).Image = 15
   If last_conditions.last_cond(0).area_of_element_no > 0 Or _
         last_conditions.last_cond(0).area_relation_no > 0 Then
   using_area_th = 8
   End If
   For i% = 0 To 3
    If conclusion_data(i%).ty = area_of_element_ Then
     using_area_th = 8
    ElseIf conclusion_data(i%).ty = area_relation_ Then
     using_area_th = 8
    End If
   Next i%
 MDIForm1.Toolbar1.Buttons(21).Image = 34
 MDIForm1.solve.Enabled = False
 Call run_
Case "markpen"
 path = App.path + "\mark.exe dingsjdingsh19751946"
  Shell (path)
Case "un_do"
Call un_do
End Select
End Sub

Private Sub Toolbar1_Change()
Draw_form.Height = screeny& - 1350 + int_w_y
Wenti_form.Height = screeny& - 1350 + int_w_y
 Draw_form.Picture1.Height = Draw_form.ScaleHeight
  Draw_form.Picture1.width = Draw_form.ScaleWidth
End Sub


Private Sub Toolbar1_DblClick()
 If finish_prove = 3 Then
 
 End If
End Sub

Private Sub total_equal_triangle_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & total_equal_triangle.Caption)
inform.Picture1.Cls
Call set_information_list(total_equal_triangle_)

End Sub


Private Sub total_picture_Click()
operator = "change_picture"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 10 '全图
old_operator = operator
 Call BitBlt(Draw_form.Picture1.hdc, 0, 0, Draw_form.Picture1.width, _
     Draw_form.Picture1.Height, Draw_form.hdc, 0, 0, vbSrcCopy)
       picture_copy = True '原图Draw_form.hdc复制到Draw_form.Picture1.hdc
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4085, "")
'Draw_form.SetFocus
'protect_munu = 1
End Sub

Private Sub turn_over_Click()
Dim i%
operator = "change_picture"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 14
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = "~."
End Sub

Private Sub turn_part_Click()
Dim i%
operator = "change_picture"
Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 4
old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2320, "")
'Draw_form.SetFocus
End Sub

Private Sub two_angle_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & two_angle.Caption)
inform.Picture1.Cls
Call set_information_list(two_angle_value_sum_)
End Sub

Private Sub two_line_value_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & two_line_value.Caption)
inform.Picture1.Cls
Call set_information_list(two_line_value_)

End Sub

Private Sub verti_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & verti.Caption)
inform.Picture1.Cls
Call set_information_list(verti_)
End Sub


Private Sub verti_mid_line_Click()
  Call remove_uncomplete_operat(old_operator)
list_type_for_draw = 2
   draw_step = 0
    Call init_draw_data
     old_operator = operator
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(2325, "")
'Draw_form.SetFocus
'protect_munu = 1
End Sub


Private Sub yizhiA_Click()
Call set_inform_list(LoadResString_(955, "") & ":" & yizhiA.Caption)
inform.Picture1.Cls
Call set_information_list(angle_value_)
End Sub
Public Sub run_()
Dim i%
'If run_type > 1 Then
'Call input_wenti_from_problem(temp_problem(0))
'End If
'temp_problem(0) = put_wenti_to_problem '记录问题
MDIForm1.Toolbar1.Buttons(19).visible = False
MDIForm1.Toolbar1.Buttons(17).visible = True
MDIForm1.Toolbar1.Buttons(18).visible = True
If finish_prove = 3 Or finish_prove = 4 Then
If finish_prove = 3 Then
Call init_no_reduce_for_condition
End If
Call start_prove(1, 1, 1)
ElseIf finish_prove = 0 Then
MDIForm1.re_name.Enabled = False
'MDIForm1.moldy_condition.Enabled = False
'MDIForm1.moldy_conclusion.Enabled = False
If run_statue = 6 Then '12.10
  Call display_run(0)
 Exit Sub
End If
prove_or_set_dbase = False
If wenti_type = 1 Then
run_statue = 2 '12.10
If start_prove(0, 1, 0) = 1 Then
For i% = 0 To 3
 If conclusion_data(i%).no(0) > 0 Then
 Wenti_form.Picture1.CurrentY = 20 * (C_display_wenti.m_last_conclusion + 1)
  Wenti_form.Picture1.CurrentX = 0
  Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(15))
     Wenti_form.Picture1.Print LoadResString_(435, "") + "(  )"
  Wenti_form.Picture1.CurrentY = 20 * (C_display_wenti.m_last_conclusion + 1)
  Wenti_form.Picture1.CurrentX = 0
  Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(0))
     Wenti_form.Picture1.Print LoadResString_(435, "");
  Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(12))
  Wenti_form.Picture1.CurrentY = Wenti_form.Picture1.CurrentY + 2
    Wenti_form.Picture1.Print LoadResString_(3955, "\\1\\" + Chr(65 + i%));
  Wenti_form.Picture1.CurrentY = Wenti_form.Picture1.CurrentY - 2
  Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(0))
    ' Wenti_form.Picture1.Print ")"
Call set_display_string_no(conclusion_data(i%).ty, conclusion_data(i%).no(0), 0, 0)
Call arrange_display_no
  Wenti_form.Picture1.CurrentY = 20 * (C_display_wenti.m_last_conclusion + 6)
  Wenti_form.Picture1.CurrentX = 0
Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(0))
    Wenti_form.Picture1.Print LoadResString_(3960, "")
  Wenti_form.Picture1.CurrentY = 20 * (C_display_wenti.m_last_conclusion + 7)
  Wenti_form.Picture1.CurrentX = 0
Call SetTextColor(Wenti_form.Picture1.hdc, QBColor(0))
Call set_display_string(True, 0, 0, 1, True)
'Call C_display_wenti.input_m_sentences(Wenti_form.Picture1, 1, C_display_wenti.m_last_conclusion + 4, _
         C_display_wenti.m_last_input_wenti_no, 1, False, 0)
End If
Next i%
End If
Else
If conclusion_data(0).ty > 0 Then
event_statue = wait_for_prove
set_or_prove = 0
display_type = 0
Call display_run(0) '开始证明过程
ElseIf prove_times = 1 Then
Call MsgBox(LoadResString_(3965, ""), 64, "", 0, 0)
End If
End If
End If

End Sub

Public Sub un_do()
Dim i%, l%, j%
Dim tp%
Call remove_uncomplete_operat(old_operator)
If MDIForm1.Toolbar1.Buttons(21).Image = 34 Then
 Exit Sub
End If
If C_display_wenti.m_last_input_wenti_no = 1 Then '初始化
Call clear_wenti_display
Call init_conditions(0)
ElseIf C_display_wenti.m_last_input_wenti_no > 1 Then
If operate_step(C_display_wenti.m_last_input_wenti_no).last_point > _
       operate_step(C_display_wenti.m_last_input_wenti_no - 1).last_point Then
 If C_display_wenti.m_no(C_display_wenti.m_last_input_wenti_no - 1) = 12 Then
   If C_display_wenti.m_point_no(C_display_wenti.m_last_input_wenti_no - 1, 2) > _
        C_display_wenti.m_point_no(C_display_wenti.m_last_input_wenti_no - 1, 4) Then
         tp% = C_display_wenti.m_point_no(C_display_wenti.m_last_input_wenti_no - 1, 2)
   Else
         tp% = C_display_wenti.m_point_no(C_display_wenti.m_last_input_wenti_no - 1, 4)
   End If
 For i% = operate_step(C_display_wenti.m_last_input_wenti_no).last_point To tp% Step -1
    Call remove_point(i%, display, 0)
 Next i%
      C_display_wenti.Remove_wenti (C_display_wenti.m_last_input_wenti_no)
'        wenti_no% = wenti_no% - 1
          'Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, _
             1, C_display_wenti.m_last_input_wenti_no, C_display_wenti.m_last_input_wenti_no, 0, 0, 0)
              Call init_wenti(C_display_wenti.m_last_input_wenti_no)
 Else
 For i% = operate_step(C_display_wenti.m_last_input_wenti_no).last_point To operate_step(C_display_wenti.m_last_input_wenti_no - 1).last_point + 1 Step -1
 Call remove_point(i%, display, 0)
 Next i%
 End If
Else '未增加点
  'wenti_no% = wenti_no% - 1
        ' Call C_display_wenti.display_m_input_condi(Wenti_form.Picture1, _
             0, C_display_wenti.m_last_input_wenti_no, C_display_wenti.m_last_input_wenti_no, 0, 0, 0)
         Call init_wenti(C_display_wenti.m_last_input_wenti_no)
  For i% = operate_step(C_display_wenti.m_last_input_wenti_no).last_con_line To operate_step(C_display_wenti.m_last_input_wenti_no).last_con_line + 1 Step -1
    'Call draw_line(Draw_form, Con_lin(i%).data(0).data0, concl, 0)
    ' l% = line_number0(m_Con_lin(i%).data(0).data0.poi(0), m_Con_lin(i%).data(0).data0.poi(1), 0, 0, True)
     ' If l% > 0 Then
      '     'Call draw_line(Draw_form, m_lin(l%).data(0).data0, condition, 0)
      'End If
      'Call init_line0(Con_lin(i%).data(0))
  Next i%
  If last_conclusion > 0 Then
  For j% = operate_step(C_display_wenti.m_last_input_wenti_no).last_conclusion To operate_step(C_display_wenti.m_last_input_wenti_no).last_conclusion + 1
   If conclusion_data(j% - 1).ty = general_string_ Then
      If con_general_string(j% - 1).data(0).value = "" Then
       For i% = last_conditions.last_cond(1).general_string_no To 1 Step -1
        If general_string(i%).record_.conclusion_no = j% - 1 Then
         Call remove_record(general_string_, i%, 0)
        End If
       Next i%
      End If
   End If
   Next j%
  draw_wenti_no = draw_wenti_no - 1
   If operate_step(C_display_wenti.m_last_input_wenti_no).last_conclusion = 0 Then
    'Call remove_solve_problem_type
   End If
  End If
  last_conclusion = operate_step(C_display_wenti.m_last_input_wenti_no).last_conclusion
End If
End If
End Sub
Private Sub zhongxinduichen_Click()
If list_type_for_draw <> 8 Then
list_type_for_draw = 8
draw_step = 0
End If
End Sub
Private Sub zhouduicheng_Click()
If list_type_for_draw <> 7 Then
list_type_for_draw = 7
draw_step = 0
End If
End Sub

Public Sub tool_bar_check()

End Sub


Public Sub Set_inpcond()
Dim last_record As Integer
Dim i%, j%
Open App.path & "\inpcond.dat" For Random As #1 Len = Len(inpcond0) '
last_record = 0
If LOF(1) > 0 Then '
Do While EOF(1) <> True
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
End If
End Sub
Public Sub set_inform_list(t_database_name As String)
   Wenti_form.List1.visible = True
   database_name = t_database_name
End Sub


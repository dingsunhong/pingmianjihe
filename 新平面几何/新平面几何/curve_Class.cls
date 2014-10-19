VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "curve_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private curve_poi() As POINTAPI
Private last_point As Integer
Private go_down_up As Integer
Private read_point As Integer
Private lin_poi(1) As POINTAPI
Private line_long As Integer
Private circ_poi(1) As POINTAPI
Private circ_radio As Integer
Private curve_ty As Byte
Private is_complete As Boolean
Private move_point_no As Integer
Private move_step As Single
Private move_position As Single
Private time_interval As Integer
Private Sub Class_Initialize()
lin_poi(0).X = -10000
lin_poi(0).Y = -10000
lin_poi(1).X = -10000
lin_poi(1).Y = -10000
circ_poi(0).X = -10000
circ_poi(0).Y = -10000
circ_poi(1).X = -10000
circ_poi(1).Y = -10000
circ_radio = 0
move_point_no = 0
last_point = 0
curve_ty = 0
is_complete = False
read_point = -1
move_step = 0.01
time_interval = 100
End Sub
Public Sub set_move_point_no(p%)
move_point_no = p%
End Sub
Public Sub Class_Init()
Call Class_Initialize
End Sub

Private Function curve_point_condinate() As POINTAPI
 If curve_ty = 1 Then
  move_position = move_position + go_down_up
move_step = move_position / line_long
curve_point_condinate.X = lin_poi(0).X + _
   (lin_poi(1).X - lin_poi(0).X) * move_step
curve_point_condinate.Y = lin_poi(0).Y + _
   (lin_poi(1).Y - lin_poi(0).Y) * move_step
  If move_position > line_long Then
    go_down_up = -1
  ElseIf move_position < 0 Then
    go_down_up = 1
  End If
ElseIf curve_ty = 2 Then
'*****************************************
move_position = move_position + move_step
curve_point_condinate.X = circ_poi(0).X + _
     circ_radio * Cos(move_position)
curve_point_condinate.Y = circ_poi(0).Y + _
     circ_radio * Sin(move_position)
'Call set_point_coordinate(p%, t_coord)
  If move_position > 2 * PI Then
   move_position = move_position - 2 * PI
  ElseIf move_position < 0 Then
   move_position = move_position + 2 * PI
  End If
'*******************************************
ElseIf curve_ty = 3 Then
   curve_point_condinate = curve_poi(read_point)
  If read_point = 0 Then
      go_down_up = 1 '上行
        read_point = 1
  ElseIf read_point = last_point Then
      go_down_up = -1 '下行
       read_point = last_point - 1
   Else
      If go_down_up = 1 Then
         read_point = read_point + 1
      Else
         read_point = read_point - 1
      End If
   End If
End If
End Function

Public Function set_curve_poi(ByVal point_no%, X As Integer, Y As Integer, _
                               X_1 As Long, Y_1 As Long, in_curve_ty As Byte, _
                                ty As Byte) As Boolean
                                'ty=1
Dim i%, s_n%, e_n%
Dim t_coord As POINTAPI
t_coord.X = X
t_coord.Y = Y
If ty = 0 Then
  Call Class_Terminate
  Call Class_Initialize
ElseIf ty = 2 Then
  is_complete = True
ElseIf ty = 3 Then '将点移到曲线上
      read_point = 0
  If point_no% > 0 Then
     move_point_no = point_no%
  If curve_ty = 3 Then '点列曲线
       go_down_up = 1
        Call move_pucture_along_curve
  ElseIf curve_ty = 2 Then '圆
     If Abs(distance_of_two_POINTAPI(m_poi(move_point_no).data(0).data0.coordinate, circ_poi(0)) - _
           circ_radio) < 5 Then '移动点在圆上
         If Abs(m_poi(move_point_no).data(0).data0.coordinate.X - circ_poi(0).X) > 10 Then
           move_position = Atn((m_poi(move_point_no).data(0).data0.coordinate.Y - circ_poi(0).Y) / _
            (m_poi(move_point_no).data(0).data0.coordinate.X - circ_poi(0).X)) '移动点和圆心连线的仰角
         Else
           move_position = Atn((m_poi(move_point_no).data(0).data0.coordinate.X - circ_poi(0).X) / _
            (m_poi(move_point_no).data(0).data0.coordinate.Y - circ_poi(0).Y))
              If m_poi(move_point_no).data(0).data0.coordinate.Y - circ_poi(0).Y > 0 Then
                 move_position = PI / 2 - move_position
              ElseIf m_poi(move_point_no).data(0).data0.coordinate.Y - circ_poi(0).Y < 0 Then
                 move_position = -PI / 2 - move_position
              End If
         End If
         If m_poi(move_point_no).data(0).data0.coordinate.X - circ_poi(0).X < 0 Then
              move_position = move_position + PI
         End If
         set_curve_poi = True
     End If
  ElseIf curve_ty = 1 Then '直线
      If read_line1(lin_poi(0), lin_poi(1), _
                            t_coord, t_coord, 0, 0, 5, False) Then
       move_position = ((lin_poi(1).X - lin_poi(0).X) * _
                        (t_coord.X - lin_poi(0).X) + _
                        (lin_poi(1).Y - lin_poi(0).Y) * _
                          (t_coord.Y - lin_poi(0).Y)) / line_long
                   
              set_curve_poi = True
       End If
  End If
  End If
   Exit Function
 ElseIf ty = 4 Then '推动点
  If curve_ty = 2 Then
       If (X_1 - circ_poi(0).X) * _
           (t_coord.Y - circ_poi(0).Y) - _
            (Y_1 - circ_poi(0).Y) * _
             (t_coord.X - circ_poi(0).X) > 0 Then
           go_down_up = 1
           move_step = 0.01
       Else
           go_down_up = -1
           move_step = -0.01
       End If
       set_curve_poi = True
  ElseIf curve_ty = 1 Then
        If (lin_poi(1).X - lin_poi(0).X) * _
            (t_coord.X - X_1) + _
           (lin_poi(1).Y - lin_poi(0).Y) * _
            (t_coord.Y - Y_1) > 0 Then
            go_down_up = 1
        Else
            go_down_up = -1
        End If
        set_curve_poi = True
  End If
   Call move_pucture_along_curve
         Exit Function
ElseIf ty = 5 Then
 If curve_ty < 3 Then
 If Abs(t_coord.X - m_poi(move_point_no).data(0).data0.coordinate.X) < 5 And _
        Abs(t_coord.Y - m_poi(move_point_no).data(0).data0.coordinate.Y) < 5 Then
         Exit Function
 End If
 If curve_ty = 2 Then
    If (m_poi(move_point_no).data(0).data0.coordinate.X - circ_poi(0).X) * _
        (t_coord.Y - circ_poi(0).Y) - _
         (m_poi(move_point_no).data(0).data0.coordinate.Y - circ_poi(0).Y) * _
          (t_coord.X - circ_poi(0).X) > 0 Then
     If go_down_up = 1 Then
      If time_interval > 50 Then
       time_interval = time_interval - 50
     End If
     ElseIf go_down_up = -1 Then
      time_interval = time_interval + 50
     End If
    Else
    If go_down_up = -1 Then
       If time_interval > 50 Then
        time_interval = time_interval - 50
       End If
    ElseIf go_down_up = 1 Then
        time_interval = time_interval + 50
    End If
  End If
 '************************
 ElseIf curve_ty = 1 Then
 '***********************
   If (lin_poi(1).X - lin_poi(0).X) * _
   (t_coord.X - m_poi(move_point_no).data(0).data0.coordinate.X) + _
    (lin_poi(1).Y - lin_poi(0).Y) * _
     (t_coord.Y - m_poi(move_point_no).data(0).data0.coordinate.Y) > 0 Then
    If go_down_up = 1 Then
     If time_interval > 50 Then
     time_interval = time_interval - 50
     End If
    ElseIf go_down_up = -1 Then
     time_interval = time_interval + 50
    End If
  Else
   If go_down_up = 1 Then
     time_interval = time_interval + 50
   ElseIf go_down_up = -1 Then
    If time_interval > 50 Then
     time_interval = time_interval - 50
    End If
   End If
  End If
 '************************
 End If
Call move_pucture_along_curve
Exit Function
End If
End If
'********************************
'画线
If curve_ty = 0 Then
  curve_ty = in_curve_ty '设置曲线类型
End If
If curve_ty = 1 Then '直线
 If lin_poi(0).X = -10000 And lin_poi(0).Y = -10000 Then
   lin_poi(0).X = X
   lin_poi(0).Y = Y
   lin_poi(1).X = X
   lin_poi(1).Y = Y
   Call m_BPset(Draw_form, lin_poi(0), "", 7)
 ElseIf ty > 0 Then
     Draw_form.Line (lin_poi(0).X, lin_poi(0).Y)- _
                (lin_poi(1).X, lin_poi(1).Y), QBColor(7)
      lin_poi(1).X = X
      lin_poi(1).Y = Y
     Draw_form.Line (lin_poi(0).X, lin_poi(0).Y)- _
                (lin_poi(1).X, lin_poi(1).Y), QBColor(7)
     If ty = 2 Then
      Call m_BPset(Draw_form, lin_poi(1), "", 7)
      line_long = sqr((lin_poi(0).X - lin_poi(1).X) ^ 2 + _
                       (lin_poi(0).Y - lin_poi(1).Y) ^ 2)
     End If
 End If
ElseIf curve_ty = 2 Then '圆
If circ_poi(0).X = -10000 And circ_poi(0).Y = -10000 Then
   circ_poi(0).X = X
   circ_poi(0).Y = Y
   circ_poi(1).X = X
   circ_poi(1).Y = Y
Else
     Draw_form.Circle (circ_poi(0).X, circ_poi(0).Y), circ_radio, _
                  QBColor(7)
      circ_poi(1).X = X
      circ_poi(1).Y = Y
      circ_radio = sqr((circ_poi(0).X - circ_poi(1).X) ^ 2 + _
                      (circ_poi(0).Y - circ_poi(1).Y) ^ 2)
     Draw_form.Circle (circ_poi(0).X, circ_poi(0).Y), circ_radio, _
                  QBColor(7)
End If
ElseIf curve_ty = 3 Then
ReDim Preserve curve_poi(last_point)
curve_poi(last_point).X = X
curve_poi(last_point).Y = Y
If ty = 0 Or ty = 2 Then
Call m_BPset(Draw_form, curve_poi(last_point), "", 7)
End If
If last_point > 0 Then
 Draw_form.Line (curve_poi(last_point - 1).X, curve_poi(last_point - 1).Y)- _
                (curve_poi(last_point).X, curve_poi(last_point).Y), QBColor(7)
End If
If ty < 2 Then
last_point = last_point + 1
End If
End If
set_curve_poi = True
End Function

Private Sub Class_Terminate()
Dim i%
On Error GoTo termnate_mark1
If curve_ty = 3 Then
   For i% = 1 To last_point
    Draw_form.Line (curve_poi(i% - 1).X, curve_poi(i% - 1).Y)- _
                (curve_poi(i%).X, curve_poi(i%).Y), QBColor(7)
   Next i%
ElseIf curve_ty = 1 Then
ElseIf curve_ty = 2 Then
     Draw_form.Circle (circ_poi(0).X, circ_poi(0).Y), circ_radio, _
                  QBColor(7)
End If
curve_ty = 0
termnate_mark1:
Erase curve_poi
End Sub
Public Sub move_pucture_along_curve()
MDIForm1.Timer1.interval = time_interval
MDIForm1.Timer1.Enabled = True
yidian_stop = False
Do
 Do
  DoEvents
 Loop Until time_act = True '启动
   Call change_picture_from_move(move_point_no, curve_point_condinate)
Call measur_again
time_act = False
Loop Until yidian_stop = True
yidian_stop = False
MDIForm1.Timer1.Enabled = False
End Sub

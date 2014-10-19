VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Icon_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event IconDisplayChange(ty As Boolean)
Private mvarm_x(1 To 32) As Integer
Private mvarm_y(1 To 32) As Integer
Private mvarm_last_position As Integer
Private mvarm_char As String * 1
Private mvarm_is_display As Boolean
Private mvarm_up_or_down As Integer
Private mvarm_icon_point As Integer
Public icon_color
Private Sub Class_Initialize()
Call init
End Sub
Public Sub init()
Dim i%
 For i% = 1 To mvarm_last_position
 mvarm_x(i%) = -10000
 mvarm_y(i%) = -10000
 Next i%
 mvarm_last_position = 0
 mvarm_icon_point% = 0
 mvarm_is_display = False
 mvarm_char = empty_char
 Call set_m_is_display(False)
 'icon_color = 15
End Sub
Public Property Get m_x(i%) As Integer
 m_x = mvarm_x(i%)
End Property
Public Property Get m_y(i%) As Integer
 m_y = mvarm_y(i%)
End Property
Public Property Get last_position() As Integer
 last_position = mvarm_last_position
End Property
Public Property Get m_char() As String
 m_char = mvarm_char
End Property
Public Property Get m_is_display() As Boolean
 m_is_display = mvarm_is_display
End Property
Public Property Get m_up_or_down() As Integer
m_up_or_down = mvarm_up_or_down
End Property
Public Sub set_m_char(icon_char$)
If mvarm_is_display = False Then
  mvarm_char = icon_char$
End If
End Sub
Public Sub set_m_is_display(ty As Boolean)
  If mvarm_is_display <> ty And mvarm_char <> empty_char Then
     mvarm_is_display = ty
       RaiseEvent IconDisplayChange(ty) '引发事件。
  End If
End Sub
Public Sub set_m_icon(icon_char$, up_or_down As Integer, icon_point%, icon_px%, icon_py%, display_or_delete As Byte)
Dim i%
  If icon_char$ <> "" Then
  If display_or_delete = 1 Then
   icon_color = 12
  Else
   Call init
    Exit Sub
  End If
  If mvarm_icon_point% = 0 Then '未设光标
   If icon_char$ = "_" Then '
     icon_point% = -2
   ElseIf (icon_char$ = global_icon_char) And icon_point% = 0 Then
     icon_point% = -1
   End If
  ElseIf mvarm_icon_point% > 0 Then '设有点光标,修改输入时用
    If mvarm_icon_point% = icon_point% Then '同点
     If icon_char$ <> global_icon_char Then '不同符号
      Exit Sub
     End If
    Else '不同点
      Exit Sub
    End If
  Else '设有光标,输入时用
   Exit Sub
  End If
  If mvarm_icon_point% = 0 Then
  mvarm_icon_point% = icon_point% '点号
  mvarm_char = icon_char$
  mvarm_up_or_down = up_or_down
  End If
  For i% = 1 To mvarm_last_position%
  If mvarm_x(i%) = icon_px% And _
       mvarm_y(i%) = icon_py% Then
     Exit Sub
  End If
  Next i%
  mvarm_last_position% = mvarm_last_position% + 1
  mvarm_x(mvarm_last_position%) = icon_px%
  mvarm_y(mvarm_last_position%) = icon_py%
  Call set_m_is_display(True)
  End If
End Sub
Public Sub display_icon(display_or_delete As Byte)
Dim i%
  If display_or_delete = 1 Then
  icon_color = 12
  Else
  icon_color = 15
  End If
For i% = 1 To mvarm_last_position
  Wenti_form.Picture1.CurrentX = mvarm_x(i%)
  Wenti_form.Picture1.CurrentY = mvarm_y(i%)
   Call display_char(Wenti_form.Picture1, mvarm_char, mvarm_up_or_down, _
          Me.icon_color, display_or_delete, False)
Next i%
End Sub
Public Function read_icon_from_string(str$, icon$, n%) As String '从字符串中读出光标
Dim i%, j%, k%, p%
Dim t_id_string As String
n% = 50
p% = 1
Do
t_id_string = read_id_string(p%, str$, j%, i%) '读出// //
  If t_id_string <> "" Then
   If i% < n% Then
      n% = i%
      read_icon_from_string = t_id_string
   End If
         p% = j% + Len(t_id_string)
  End If
Loop Until t_id_string = "" Or read_icon_from_string = "_" Or read_icon_from_string = global_icon_char
If read_icon_from_string = "_" Or read_icon_from_string = global_icon_char Then
   Exit Function
ElseIf read_icon_from_string <> "" Then
   If Len(str$) >= p% + Len(read_icon_from_string) Then
        If Mid$(str$, p% + Len(read_icon_from_string), 1) = "_" Then
                icon$ = "_"
        Else
                icon$ = global_icon_char
        End If
   Else
        icon$ = global_icon_char
   End If
End If
End Function
Public Sub Terminate()
Me.Terminate
End Sub
Public Sub close_icon()
Call display_icon(0)
Call init
End Sub

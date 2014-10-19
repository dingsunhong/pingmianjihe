VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "icons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_icons As New Collection
Public m_is_display As Boolean
Public Sub set_new_icon(icon_char$, up_or_down As Integer, display_or_delete As Byte)
Dim i%, m_x%, m_y%
Dim m_icon As New Icon_Class
m_x% = Wenti_form.Picture1.CurrentX
m_y% = Wenti_form.Picture1.CurrentY
For i% = 1 To Me.m_icons.Count
 If Me.m_icons.item(i%).m_x% = m_x And _
     Me.m_icons.item(i%).m_y% = m_y Then
      Exit Sub
 End If
Next i%
Me.m_icons.Add m_icon
m_is_display = True
MDIForm1.Timer1.Enabled = True
Call Me.m_icons.item(m_icons.Count).set_m_icon(icon_char$, up_or_down, m_x%, m_y%, display_or_delete)
End Sub
Public Sub delete_icons()
 m_is_display = False
 Set Me.m_icons = Nothing
 Set Me.m_icons = New Collection
End Sub
Public Sub set_icons_is_display(ty As Boolean)
Dim i%
m_is_display = ty
For i% = 1 To Me.m_icons.Count
Me.m_icons.item(i%).set_m_is_display (ty)
Next i%
End Sub
Public Sub display_icons(dis_or_delete As Byte)
For i% = 1 To Me.m_icons.Count
Me.m_icons.item(i%).display_icon (dis_or_delete)
Next i%
End Sub

Private Sub Class_Initialize()
Call delete_icons
End Sub

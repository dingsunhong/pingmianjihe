VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "coordinate_system0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private name As String
Private direction As Integer
Private is_display_origin As Boolean
Private origin_coord As POINTAPI
Private start_point As POINTAPI
Private end_point As POINTAPI
Private unit As Integer
Private axi_length As Integer
Private color As Integer
Private ob As Object
Public Sub draw_coordinate()
Dim half_l As Integer
half_l = axi_length / 2

End Sub


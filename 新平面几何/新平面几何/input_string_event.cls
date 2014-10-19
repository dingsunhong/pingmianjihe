VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "input_string_event"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Event InputComplete(w_n As Integer)
 Event IconDisplayChange(w_n As Integer, ty As Boolean)
Public Sub set_input_complete(w_n%)
 RaiseEvent InputComplete(w_n%)
End Sub

Public Sub set_input_icon(w_n%, ty As Boolean)
 RaiseEvent IconDisplayChange(w_n%, ty)
End Sub

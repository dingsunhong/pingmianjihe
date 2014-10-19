VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "condition_tree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public m_condition As New Collection
Public condition_head As Integer
Public condition_tail As Integer
Private Sub Class_Initialize()
Set m_condition = Nothing
Set m_condition = New Collection
condition_head = 0
condition_tail = 0
End Sub
Public Sub initial()
Call Class_Initialize
End Sub

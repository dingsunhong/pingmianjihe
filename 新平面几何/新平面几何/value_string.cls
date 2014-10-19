VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "value_string"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Global Const polynomial = 0
' Global Const squre_root_ = 1
' Global Const ln_ = 2
' Global Const exp_ = 3
' Global Const sin_ = 4
' Global Const cos_ = 5
' Global Const tan_ = 6
' Global Const ctan_ = 7
' Global Const Asin_ = 8
' Global Const Acos_ = 9
' Global Const Atan_ = 10
' Global Const Actan_ = 11
'
'**************************************************************************
Option Explicit
Private m_value As String
Public m_ty As Byte
Private m_para As String
Private m_value_for_int(1) As Integer
Private m_value_for_float As Single
Private Sub Class_Initialize()
m_value = ""
m_ty = 0
m_para = ""
End Sub
Public Function value() As String
 value = m_value
End Function
Public Sub set_value(value As String)
 m_value = value
End Sub
Public Function para() As String
 para = m_para
End Function
Public Sub set_para(para As String)
 m_para = para
End Sub
'Public Function ty() As Byte
' ty = m_ty
'End Function
'Public Sub set_ty(value As Byte)
' m_ty = ty
'End Sub


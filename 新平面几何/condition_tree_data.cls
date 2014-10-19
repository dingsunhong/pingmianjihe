VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "condition_tree_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mavrm_ty As Byte
Private mavrm_no As Integer
Private conditions_ty(8)  As Byte
Private conditions_no(8)  As Integer
Public conclusion_no As Byte
Private father_condition_no(1) As Integer
Private temp_father(1) As Integer
Public pre_no As Integer
Public next_no As Integer



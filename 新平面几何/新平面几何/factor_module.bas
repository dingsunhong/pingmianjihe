Attribute VB_Name = "factor_module"
Public Sub copy_factor_to_factor(fa_1 As factor_type, fa_2 As factor_type)
Call copy_factor0_to_factor0(fa_1.data(0), fa_2.data(0))
Call copy_factor0_to_factor0(fa_1.data(1), fa_2.data(1))
End Sub

Public Sub copy_factor0_to_factor0(fa0_1 As factor0_type, fa0_2 As factor0_type)
Dim i%
fa0_2.last_factor = fa0_1.last_factor
fa0_2.para = fa0_1.para
For i% = 1 To fa0_1.last_factor
ReDim Preserve fa0_2.factor(i%) As String
 fa0_2.factor(i%) = fa0_1.factor(i%)
ReDim Preserve fa0_1.order(i%) As Integer
 fa0_2.order(i%) = fa0_1.order(i%)
Next i%
End Sub


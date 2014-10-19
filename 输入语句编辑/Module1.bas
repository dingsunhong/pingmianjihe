Attribute VB_Name = "Module1"
'文件中的结构调整
Type taboo_type
taboo_relation(1) As Integer
ty As Integer
End Type
Type inpcond_type
no As Integer
ty As Byte
'chinese_inpcond As String * 100
inpcond(4) As String * 128
'chinese_and_fogrein(10, 1) As Integer
relation(1, 1) As Integer
taboo(7) As taboo_type
End Type
'***************************************
'程序中的结构
Type inpcond0_type
no As Integer
ty As Byte
'chinese_inpcond As String * 100
inpcond As String * 128
'chinese_and_fogrein(10, 1) As Integer
relation(1, 1) As Integer
taboo(7) As taboo_type
End Type
'**************************************************
Type inpcond1_type
no As Integer
ty As Byte
'chinese_inpcond As String * 100
inpcond(4) As String * 256
'chinese_and_fogrein(10, 1) As Integer
relation(1, 1) As Integer
taboo(7) As taboo_type
End Type
Type inpcond10_type
no As Integer
ty As Byte
'chinese_inpcond As String * 100
'inpcond As String * 128
'chinese_and_fogrein(10, 1) As Integer
relation(1, 1) As Integer
taboo(7) As taboo_type
End Type

'***************************************
Global inpcond0 As inpcond_type
Global inpcond10 As inpcond1_type
Global last_record As Integer


Public Sub get_data(data_no%)
  Get #1, data_no%, inpcond10
    Form1.Text1.Text = inpcond10.no
    Form1.Text2.Text = inpcond10.inpcond(0)
    Form1.Text3.Text = inpcond10.inpcond(1)
    Form1.Text4.Text = inpcond10.inpcond(2)
    Form1.Text5.Text = inpcond10.inpcond(3)
End Sub
Public Sub set_data(data_no%)
    inpcond10.inpcond(0) = Form1.Text2.Text
    inpcond10.inpcond(1) = Form1.Text3.Text
    inpcond10.inpcond(2) = Form1.Text4.Text
    inpcond10.inpcond(3) = Form1.Text5.Text
   Put #1, data_no%, inpcond10
End Sub

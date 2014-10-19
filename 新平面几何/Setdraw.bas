Attribute VB_Name = "setdraw"
Global temp_line_width As Byte
Global temp_condition_color As Byte
Global temp_conclusion_color As Byte
Global temp_fill_color As Byte
Global line_width  As Byte
Global condition_color As Byte
Global conclusion_color As Byte
Global fill_color  As Byte


Public Sub draw_conclusion()
widthform.Picture6.DrawWidth = temp_line_width
widthform.Picture6.Line (20, 20)-(80, 20), QBColor(temp_conclusion_color)
End Sub

Public Sub linewidth()
widthform.Picture4.DrawWidth = temp_line_width
widthform.Picture4.Line (20, 20)-(80, 20), QBColor(0)

End Sub

Public Sub draw_sample()
widthform.Picture3.Cls
widthform.Picture3.DrawWidth = temp_line_width
widthform.Picture3.Circle (130, 80), 70, QBColor(temp_condition_color)
widthform.Picture3.Line (130, 10)-(191, 115), QBColor(temp_condition_color)
widthform.Picture3.Line (69, 115)-(191, 115), QBColor(temp_condition_color)
widthform.Picture3.Line (130, 10)-(69, 115), QBColor(temp_condition_color)
widthform.Picture3.Line (130, 10)-(100, 143), QBColor(temp_conclusion_color)
widthform.Picture3.Line (69, 115)-(100, 143), QBColor(temp_conclusion_color)
widthform.Picture3.Line (191, 115)-(100, 143), QBColor(temp_conclusion_color)
If temp_line_width < 3 Then
 widthform.Picture3.PaintPicture Draw_form.ImageList1.ListImages(Asc("A") - 60).Picture, 128, 8, 16, 18, 0, 0, 16, 18, &H990066
 widthform.Picture3.PaintPicture Draw_form.ImageList1.ListImages(Asc("B") - 60).Picture, 67, 113, 16, 18, 0, 0, 16, 18, &H990066
 widthform.Picture3.PaintPicture Draw_form.ImageList1.ListImages(Asc("C") - 60).Picture, 189, 113, 16, 18, 0, 0, 16, 18, &H990066
 widthform.Picture3.PaintPicture Draw_form.ImageList1.ListImages(Asc("D") - 60).Picture, 98, 141, 16, 18, 0, 0, 16, 18, &H990066
Else
 widthform.Picture3.PaintPicture Draw_form.ImageList3.ListImages(Asc("A") - 60).Picture, 124, 4, 16, 18, 0, 0, 16, 18, &H990066
 widthform.Picture3.PaintPicture Draw_form.ImageList3.ListImages(Asc("B") - 60).Picture, 63, 109, 32, 32, 0, 0, 32, 32, &H990066
 widthform.Picture3.PaintPicture Draw_form.ImageList3.ListImages(Asc("C") - 60).Picture, 185, 109, 32, 32, 0, 0, 32, 32, &H990066
 widthform.Picture3.PaintPicture Draw_form.ImageList3.ListImages(Asc("D") - 60).Picture, 94, 137, 32, 32, 0, 0, 32, 32, &H990066
End If
End Sub

Public Sub draw_condition_color()
widthform.Picture5.DrawWidth = temp_line_width
widthform.Picture5.Line (20, 20)-(80, 20), QBColor(temp_condition_color)
End Sub
Public Sub draw_fill_color()
widthform.Picture7.DrawWidth = temp_line_width
widthform.Picture7.Line (20, 20)-(80, 20), _
   QBColor(temp_fill_color), BF
End Sub



Public Sub init_set()
'Call init_color
Call draw_sample
Call draw_conclusion
Call draw_condition_color
Call draw_fill_color
Call linewidth
Call draw_arow
'widthform.label1.BackColor = QBColor(15)
'widthform.Label2.BackColor = QBColor(15)
'widthform.Label3.BackColor = QBColor(15)
'widthform.Label4.BackColor = QBColor(15)
'set_statue = 0

'widthform.Picture2.PaintPicture widthform.Image1, 10, 0
End Sub

Public Sub draw_arow()
widthform.Picture2.Cls
Select Case temp_line_width
Case 1
widthform.Picture2.PaintPicture widthform.Image1, 10, 0
Case 2
widthform.Picture2.PaintPicture widthform.Image1, 10, 5
Case 3
widthform.Picture2.PaintPicture widthform.Image1, 10, 10
Case 4
widthform.Picture2.PaintPicture widthform.Image1, 10, 16
Case 5
widthform.Picture2.PaintPicture widthform.Image1, 10, 24
End Select
End Sub

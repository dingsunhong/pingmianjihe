Attribute VB_Name = "menu_module"
Public Sub set_menu_for_set_function_data0()
MDIForm1.image_of_function.visible = False
MDIForm1.length.Caption = LoadResString_(4165, "") + LoadResString_(4140, "") '自变量设为长度"
MDIForm1.distance_p_line.Caption = LoadResString_(4165, "") + LoadResString_(4145, "") ' "自变量设为点到直线的距离"
MDIForm1.eara.Caption = LoadResString_(4165, "") + LoadResString_(4150, "") '"自变量设为多边形面积"
MDIForm1.angle.Caption = LoadResString_(4165, "") + LoadResString_(4155, "") '"自变量设为角度"
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4160, "\\" + LoadResString_(4170, ""))  ' "设函数的自变量"
End Sub
Public Sub set_menu_for_set_function_data1()
MDIForm1.image_of_function.visible = False
MDIForm1.length.Caption = LoadResString_(4170, "") + LoadResString_(4140, "") '
MDIForm1.distance_p_line.Caption = LoadResString_(4170, "") + LoadResString_(4145, "") ' "因变量设为点到直线的距离"
MDIForm1.eara.Caption = LoadResString_(4170, "") + LoadResString_(4150, "") '"因变量设为多边形面积"
MDIForm1.angle.Caption = LoadResString_(4170, "") + LoadResString_(4155, "") '"因变量设为角度"
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4160, "\\1\\" + LoadResString_(4170, "")) ' "设函数的因变量"
End Sub
Public Sub recove_set_menu_for_set_function_data()
MDIForm1.image_of_function.visible = True
MDIForm1.length.Caption = LoadResString_(855, "") '"测量长度"
MDIForm1.distance_p_line.Caption = LoadResString_(890, "") ' "测量点到直线的距离"
MDIForm1.eara.Caption = LoadResString_(895, "") '"测量多边形面积"
MDIForm1.angle.Caption = LoadResString_(900, "") '"测量角度"
End Sub


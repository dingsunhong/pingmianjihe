Attribute VB_Name = "menu_module"
Public Sub set_menu_for_set_function_data0()
MDIForm1.image_of_function.visible = False
MDIForm1.length.Caption = LoadResString_(4165, "") + LoadResString_(4140, "") '�Ա�����Ϊ����"
MDIForm1.distance_p_line.Caption = LoadResString_(4165, "") + LoadResString_(4145, "") ' "�Ա�����Ϊ�㵽ֱ�ߵľ���"
MDIForm1.eara.Caption = LoadResString_(4165, "") + LoadResString_(4150, "") '"�Ա�����Ϊ��������"
MDIForm1.angle.Caption = LoadResString_(4165, "") + LoadResString_(4155, "") '"�Ա�����Ϊ�Ƕ�"
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4160, "\\" + LoadResString_(4170, ""))  ' "�躯�����Ա���"
End Sub
Public Sub set_menu_for_set_function_data1()
MDIForm1.image_of_function.visible = False
MDIForm1.length.Caption = LoadResString_(4170, "") + LoadResString_(4140, "") '
MDIForm1.distance_p_line.Caption = LoadResString_(4170, "") + LoadResString_(4145, "") ' "�������Ϊ�㵽ֱ�ߵľ���"
MDIForm1.eara.Caption = LoadResString_(4170, "") + LoadResString_(4150, "") '"�������Ϊ��������"
MDIForm1.angle.Caption = LoadResString_(4170, "") + LoadResString_(4155, "") '"�������Ϊ�Ƕ�"
MDIForm1.StatusBar1.Panels(1).text = LoadResString_(4160, "\\1\\" + LoadResString_(4170, "")) ' "�躯���������"
End Sub
Public Sub recove_set_menu_for_set_function_data()
MDIForm1.image_of_function.visible = True
MDIForm1.length.Caption = LoadResString_(855, "") '"��������"
MDIForm1.distance_p_line.Caption = LoadResString_(890, "") ' "�����㵽ֱ�ߵľ���"
MDIForm1.eara.Caption = LoadResString_(895, "") '"������������"
MDIForm1.angle.Caption = LoadResString_(900, "") '"�����Ƕ�"
End Sub


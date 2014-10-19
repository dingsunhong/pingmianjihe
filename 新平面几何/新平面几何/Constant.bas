Attribute VB_Name = "const"
Option Explicit
'**********************************************************
'ÓÃÓÚ¶ÁÓ²ÅÌÐòÁÐºÅ
'Public Declare Function HDSerialNumRead Lib "HDSerialNumRead.dll" () As String

'Public Type OSVERSIONINFO
'        dwOSVersionInfoSize As Long
'        dwMajorVersion As Long
'        dwMinorVersion As Long
'        dwBuildNumber As Long
'        dwPlatformId As Long
'        szCSDVersion As String * 128      '  Maintenance string for PSS usage
'End Type
'Public Const VER_PLATFORM_WIN32s = 0
'Public Const VER_PLATFORM_WIN32_WINDOWS = 1
'Public Const VER_PLATFORM_WIN32_NT = 2
'Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'***************************************************************************************************************************
Type link_sock_type
ip As String
name As String
End Type
'**************************************************************************************************
Type v_string
 coord(1) As String
End Type
'Global protect_munu As Integer '±£»¤²Ëµ¥,µã»÷²Ëµ¥,²»ÈÃÊó±ê,Ó¡µ½draw_form,=1,µã»÷Êó±ê£¬=0£¬=0,»­Í¼
'Global protect_munu_ As Integer
Global Mdiform1_caption As String
Global display_information_string(0 To 100) As String
Global Const verti_mark_meas = 8
Global width_set_statue As Integer
Global StrOpenFile As String
Global LoginSucceeded As Boolean
Global is_temp_inform_dll_open As Boolean
Global data_no% '¼ÆÊ±Æ÷¼ÇÂ¼´ÎÊý
Global tangent_circle_type As Byte '»­ÇÐÔ²µÄÀàÐÍ
Global input_coord As POINTAPI 'Êó±êÊäÈëµÄµãµÄ×ø±ê
Global t_coord As POINTAPI
Global t_coord1 As POINTAPI
Global t_coord2 As POINTAPI
Global pointapi0 As POINTAPI '¹©º¯Êý²ÎÊýÓÃ
Global last_node_index As Integer 'Ê÷½á¹¹½áµãÊý
Global path_and_file As String '´ò¿ªÎÊÌâµÄÂ·¾¶
Global save_statue As Byte 'ÎÊÌâµÄ±£´æ×³Ì¬
Global wenti_no_% 'wenti_no% ¸±±¾
'Global old_wenti_no_% 'old_wenti_no% ¸±±¾
Global picture_copy As Boolean 'ÏÔÊ¾ÐÅÏ¢µÄÍ¼Ê¾×´Ì¬
Global chose_w_no% 'Ñ¡ÖÐµÄÎÊÌâµÄÓï¾äºÅ
Global is_new_result As Boolean '¼ÌÐøÍÆÀíÊÇ·ñÓÐÐÂ½á¹û,(Î´ÓÃ????)
Global exam_wenti_name() As String * 20 '´æ·ÅÀýÌâÄ¿Â¼
Global circle_with_center As Boolean 'ÊÇ·ñ»­ÓÐÐÄÔ²
Global write_wenti As Boolean
Global change_pic As Integer
Global re_name_ty As Byte
Global WindowsDirectory As String
Global SystemDirectory As String
Global database_name As String
Global protect_file(2) As String '0 The Mircosolft_sound µÄ¸±±¾,1 =d.ll
Global Const dll_code = "MZPÿÿ@"
Global yes_Id As Boolean
Global set_change_type_ As Boolean
Global system_vision As Integer
'Global temp_th_chose(-5 To 180) As Byte
Global trajectory(1000) As POINTAPI
Global Const PI = 3.14159265358979
Public Const OF_READ = &H0
Global is_uselly_para_for_angle As Boolean 'angle3_valueÖÐÏµÊýÖ»ÄÜµÈÓÚ0 1,2
Global is_uselly_degree_for_angle As Boolean
'Global turn_over_type As Integer
Global chose_total_theorem  As Boolean
Global SSTab1_name_type As Byte  '0 ÄæÏòÍÆÀí  =1 Êý¾Ý¿â = 2 ²úÉúÊý¾Ý¿â
'*****************
Global error_of_wenti As Byte
Global error_condition_no As Integer
Global solve_problem_type As Byte '¼ÇÂ¼0,ÇóÖ¤,1.Çó3Ñ¡Ôñ
Global paral_or_verti As Integer 'ÊäÈë¹ý³ÌÖÐ¼ÇÂ¼Æ½ÐÐ´¹Ö±
Type operate_record
last_point As Integer
last_con_line As Integer
last_con_circle As Integer
last_conclusion As Integer
End Type
Global operate_step(30) As operate_record
Global computer_id_ As String * 8
Type protect_data_type
  id As String * 8
  computer_id As String * 8
  pass_word As String * 12
  input_pass_word_time As Integer
  install_date As String * 8
  used_time As Integer
  install_statue As String * 1 '"S",µÚÒ»´Î°²×° "F",¹ýÆÚ,"T",Ö¤Êµ
  serial_no As String
  pass_word_for_teacher As String * 5
End Type
Global protect_data As protect_data_type
'******************
Global last_combine_length_of_polygon_with_line_value(1) As Integer
Global last_combine_length_of_polygon_with_two_line_value(1) As Integer
Global last_combine_length_of_polygon_with_line3_value(1) As Integer
'*****************
'ÓÃÓÚÈ«µÈÏàËÆ
'**********************************
Type temp_triangle_data_type
 poi(2) As Integer
  angle(2) As Integer
   l_v(2) As Integer
     no As Integer
      direction As Integer
       is_contain_p As Boolean
End Type
Type temp_triangle_type
 data(50) As temp_triangle_data_type
      last_T As Byte
End Type
'********************
Global conclu_ty As Integer
'*************************
Type condition_no_type
ty As Byte
no As Integer
End Type
'************************
'
'*************************
Type add_point_for_two_line_type
 poi As Integer 'poi=0 Ïà½»
  s_poi(1) As Integer 'Æðµã
   paral_or_verti(1) As Byte
    line_no(1) As Integer
     index As Integer
End Type
Type add_point_for_two_circle_type
circ(1) As Integer
    index As Integer
End Type
Type add_point_for_line_circle_type
 poi As Integer 'poi=0 Ïà½»
  paral_or_verti As Byte
   line_no As Integer
    circ As Integer
     index As Integer
End Type
Type add_point_for_mid_point_type
 line_no As Integer
  poi(2) As Integer
   index As Integer
End Type
Type add_point_for_eline_type
poi(2) As Integer
 line_no As Integer
  te As Byte
   index As Integer
End Type
Type wait_for_add_point_type
ty As Byte
next_no As Integer
 para(0 To 8) As Integer
  last_para As Integer
End Type
'*************************************
Global Const interset_point_three_line = 2
Global Const aid_point_for_circle1 = 3
'********************************
Global add_aid_point_for_two_line_() As add_point_for_two_line_type
Global last_add_aid_point_for_two_line As Integer
Global add_aid_point_for_two_line0 As add_point_for_two_line_type
Global add_aid_point_for_two_circle_() As add_point_for_two_circle_type
Global last_add_aid_point_for_two_circle As Integer
Global add_aid_point_for_two_circle0 As add_point_for_two_circle_type
Global add_aid_point_for_line_circle_() As add_point_for_line_circle_type
Global last_add_aid_point_for_line_circle As Integer
Global add_aid_point_for_line_circle0 As add_point_for_line_circle_type
Global add_aid_point_for_mid_point_() As add_point_for_mid_point_type
Global last_add_aid_point_for_mid_point As Integer
Global add_aid_point_for_mid_point0 As add_point_for_mid_point_type
Global add_aid_point_for_eline_() As add_point_for_eline_type
Global last_add_aid_point_for_eline As Integer
Global add_aid_point_for_eline0 As add_point_for_eline_type
Global Const empty_char = " "
Global select_wenti_no%
Global temp_polygon As polygon
Global set_measure_no%
Global Const co0 = &HFF8
Global Const co1 = &H8080FF
Global Const co2 = &HFF0
Global Const co3 = &HC00
Global Const change_L = "control_L"
'Global Const OF_EXSIT = &O4000
Global display_wenti_h_position%
'Global display_wenti_v_position%
Global cond_no() As condition_no_type
'Global temp_total_condition  As Integer
Global start_no%
Global line3_value_conclusion As Byte
Global line3_angle_conclusion As Byte
Global area_of_triangle_conclusion As Byte
Global start_type As Integer
Global char As String
'¹©ÊäÈëÓÃ
Global reduce_level As Byte
Global reduce_level0 As Byte
Global ge_reduce_level As Byte 'general_stringµÄÍÆÀíÉî¶È£¬ÓÉÌõ¼þ¾ö¶¨
Global contro_process As Integer
Global finish_prove As Byte
Global text_num1%
Global text_num2%
Global text_num!
Global wenti_form_title As String
'Global new_result_from_add As Boolean
Global pro_no%  '±£»¤°æÈ¨
Global pro_no1%
Global run_statue As Byte
Global chapter_no As Integer 'Ñ¡½ø¶È
Global wenti_form_treeview_visible As Boolean
Global wenti_form_picture_visible As Boolean
Global inform_treeview_visible As Boolean
Global inform_picture_visible As Boolean
Global temp_th_ch_51  As Byte
Global temp_th_ch_52  As Byte
Type taboo_type
taboo_relation(1) As Integer
ty As Integer
End Type
'**********************************************
'ÎÄ¼þÖÐµÄ½á¹¹µ÷Õû
Global value_for_draw(8) As Single
Type inpcond_type
no As Integer
ty As Byte
'chinese_inpcond As String * 100
inpcond(4) As String * 256
'chinese_and_fogrein(10, 1) As Integer
relation(1, 1) As Integer
taboo(7) As taboo_type
End Type
'***************************************
'³ÌÐòÖÐµÄ½á¹¹
Type inpcond0_type
no As Integer
ty As Byte
'chinese_inpcond As String * 100
inpcond As String * 128
'chinese_and_fogrein(10, 1) As Integer
relation(1, 1) As Integer
taboo(7) As taboo_type
End Type

'***************************************
Global inpcond0 As inpcond_type
Type line_from_two_point_data
 n(1) As Integer
  v_poi(1) As Integer
   v_n(1) As Integer
    line_no As Integer
     index As Integer
      no_reduce As Byte
       value As String
        v_line_value_no As Integer
         v_value As String
          dir As Integer
End Type
Global line_from_two_point_data_0 As line_from_two_point_data
Type line_from_two_point
poi(1) As Integer
  data(8) As line_from_two_point_data
End Type
Global Dline_from_two_point_data As line_from_two_point_data
Global Dtwo_point_line() As line_from_two_point
'Global last_conditions.last_cond(1).line_no_from_two_point(1) As Integer
'Global last_conditions.last_cond(1).line_no_from_two_point_for_aid(1) As Integer
'Type chapter_type
'text As String
'no As Integer
'End Type
'Global chapter(110) As chapter_type
'Global last_chapter As Integer
' 0ÊäÈë£¬1×Ô¶¯ ÕûÌåÏÔÊ¾ 2 ×Ô¶¯·Ö²½ÏÔÊ¾£¬ 3 ½»»¥ 4¡£¼Ó¸¨ÖúÏß
Global ruler_display As Boolean
'Global total_condition As Integer
Global last_condition_total_condition As Integer
Global old_total_condition As Integer
Global last_add_condition As Integer
Global choose_point As Integer
Global prove_times As Byte
Global last_th_choose As Integer
Global display_inform As Byte '¼ÇÂ¼ÊÖ¹¤Ö¤Ìâ ÊÇ·ñÖ±½ÓÍÆÀí
Global int_w_y As Integer '´°¿ÚµÄ¸ß¶È
Global arrange_window_type  As Byte '´°¿ÚÅÅÁÐ·½Ê½
Global prove_or_set_dbase As Boolean
Global input_text_statue As Boolean ' ¼ÇÂ¼mdiform1.text2µÄÊäÈë
Global input_text_finish As Boolean
Global last_conclusion As Integer
Global wenti_type As Byte '¼ÇÂ¼Ö¤Ã÷0£¬Ñ¡Ôñ1
'Global wenti_type0 As Byte '0 ÇóÖ¤ 1 ½â
Global top0%, left0%
Global old_operator As String
Global list_type_for_draw As Integer
Global is_first_move As Boolean
Global set_change_fig As Byte '±íÊ¾ÊÇ·ñÑ¡ÖÐ±ä»»Í¼ÀàÐÍ
Global display_prove_proccess As Boolean
Global operat_is_acting As Boolean
Global center_p As POINTAPI
Global yidian_type As Byte
Global yidian_stop As Boolean
Global time_act As Boolean
Global draw_time_act As Boolean
Global set_or_prove As Byte
Global prove_type As Byte
'Global draw_step As Integer
Global draw_operate As Boolean
Global draw_point_no  As Integer
Global problem_type As Boolean
'Global c_display_wenti.m_last_conclusion As Integer
'Global top_sentence_no As Integer
'Global ratio_for_measur As Single
Global screenx&, screeny&
Global printrx&, printry&
Global cond_type As Byte
Global old_point As Integer
Global last_constant As Integer
'ÍÆÀíÖÐÉè¶¨µÄ³£Á¿
Type TH_chose_type
chose As Byte ' 0 Ñ¡ÖÐ
used As Byte
chapter As Integer
text As String
TH_name As String
End Type
Global th_chose(-6 To 180) As TH_chose_type
'Global point_inform(26) As String
Type point_pair_for_similar_data_type
triA(1) As Integer
direction(1) As Integer
point_pair_no(2) As Integer
is_proved As Byte
End Type
Global point_pair_for_similar_data_0 As point_pair_for_similar_data_type
Type point_pair_for_similar_type
data(8) As point_pair_for_similar_data_type
End Type
Global draw_or_prove As Byte
Global point_pair_for_similar() As point_pair_for_similar_type
'Global last_point_pair_for_similar As Integer
'±ÈÀýÏß¶ÎÖÐºòÑ¡ÏàËÆ
'ÆÁÄ»ºÍ´òÓ¡»úµÄ·Ö±ç
'ÓÃÓÚ²âÁ¿Ïß¶Î³¤
Type length_type
 string_no As Byte
  poi(1) As Integer
   equal_line As Integer
   len As Single
    len0 As Single
End Type
Global length_(20) As length_type
Global last_length As Integer
'*************************************************
'µãµ½Ïß¶ÎµÄ¾àÀë
Type length_point_to_line_type
 string_no As Byte
  poi(2) As Integer
   line_no As Integer
   len As Long
    len0 As Long
End Type
Global length_point_to_line(20) As length_point_to_line_type
Global last_length_point_to_line As Integer
'*********************************************
Type unkown_element_type
char As String
conclusion_no As Integer
End Type
Global unkown_element(1 To 8) As unkown_element_type
'****************************************
Type view_point_type
poi As Integer
old_coordinate As POINTAPI
End Type
Global view_point(1 To 8) As view_point_type
'¶à±ßÐÎ
Type line0_type
poi(1) As Integer
line_no As Integer
index As Integer
old_index As Integer
End Type
Global line0() As line0_type
'Global last_conditions.last_cond(1).line_no0 As Integer
'Global last_conditions.last_cond(1).line_no0_for_aid As Integer
Type polygon
v(15) As Integer
line_no(15) As Integer
coord(15) As POINTAPI
total_v As Byte
center As POINTAPI
coord_center As POINTAPI
direction As Boolean
is_e_polygon As Boolean
End Type
Global polygon_data_0 As polygon
'¶à±ßÐÎÃæ»ý
Global poly() As polygon
'Global last_poly As Byte
Type Area_polygon_type
 string_no As Byte
  p As polygon
   Area As Long
    Area0 As Long
End Type
Global Area_polygon(20) As Area_polygon_type
Global last_Area_polygon As Byte
'²âÁ¿½Ç
Type edit_char_type
ch As String
pos As POINTAPI
End Type
Global E_char() As edit_char_type
Global last_edit_char As Integer
Type line_for_change0_type
poi(1) As Integer
in_point(10) As Integer
in_poi_coord(1 To 10) As POINTAPI
coord(1) As POINTAPI
center(1) As POINTAPI
End Type
Type line_for_change_type
line_no(1) As line_for_change0_type
move As POINTAPI
rote_angle As Single
direction As Integer
similar_ratio As Single
End Type
Global line_for_change As line_for_change_type
Type polygon_for_change_type
p(1) As polygon
move As POINTAPI
rote_angle As Single
direction As Integer
similar_ratio As Single
End Type
Type symmetry_line_type
 p(1) As POINTAPI
End Type
Global symmetry_line As symmetry_line_type
Global symmetry_point As POINTAPI
'Type vector_data0_type
'poi(1) As Integer
'n(1) As Integer
'line_no As Integer
'dir As Integer
'value As String
'value_no As Integer
'End Type
'Type vector_type
'data(8) As vector_data0_type
'End Type
'Global Dvector() As vector_type
Global dot_line(1 To 16) As line_data0_type
'Global last_dot_line As Integer
Type angle_value_for_measur_type
 string_no As Byte
 poi(2) As Integer
  angle As Integer
  eangle As Integer
  value As String
End Type
Global angle_value_for_measur(20) As angle_value_for_measur_type
Global last_angle_value_for_measur As Byte
Global Measur_string(50) As String
Global last_measur_string As Byte
'µãµÄ×ù±ê
'»­Í¼Ê±ÏÔÊ¾½»µãµÄÏß
Type Tsysinfor
   WindowsProductID As String
   MACAddress As String
   BiosVersion As String
   BiosSerialNum As String
   HDSerialNum As String
   HDVolSerialNum As String
   SysHDSize As String
 End Type
'Declare Function HDSerialN _
        Lib "HDSerialNum" () As Tsysinfor
Public Const OFS_MAXPATHNAME = 128
Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
        
Type FileTime 'ÎÄ¼þÊ±¼ä
     dwLowDateTime As Long
     dwHighDateTime As Long
End Type
Type FILE_TIME 'ÎÄ¼þ
     'dwFileAttributes As Long
     ftCreationTime As FileTime
     ftLastAccessTime As FileTime
     ftLastWriteTime As FileTime
     'dwVolumeSerialNumber As Long
     'nFileSizeHigh As Long
     'nFileSizeLow As Long
     'nNumberOfLinks As Long
     'nFileIndexHigh As Long
     'nFileIndexLow As Long
End Type
Type SYS_TIME
Year As Integer
Month As Integer
Dayofweek As Integer
Day As Integer
Hour As Integer
Minite As Integer
Second As Integer
Millsecond As Integer
End Type
Type SysInform_records
WindowsProductID As String
MACAddress As String
BiosVersion As String
BiosSerialNum As String
HDVolSerialNum As String
SysHDSize As String
HDSerialNumber As String
End Type
Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Type HDFingerPrint
  HDModelNumber As String * 41
  HDFirmwareRev As String * 9
  HDSerialNumber As String * 21
  HDCapacity(2) As Long
End Type
Global protect_file_time(2) As FILE_TIME
Global protect_sys_time(3) As SYS_TIME
Global lpReOpenBuff As OFSTRUCT
Global FileHandle(1) As Long
Public Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type
Public Declare Function PolyBezier Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
Public Declare Function PolyBezierTo Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FileTime, lpSystemTime As SYS_TIME) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYS_TIME, lpFileTime As FileTime) As Long
Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
    lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
    '´ò¿ªÎÄ¼þ
'Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
   '¹Ø±Õ
Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, _
      lpCreationTime As FileTime, lpLastAccessTime As FileTime, _
                 lpLastWriteTime As FileTime) As Long
Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, _
           lpCreationTime As FileTime, lpLastAccessTime As FileTime, _
                  lpLastWriteTime As FileTime) As Long
Public Declare Function HDSerialNumRead Lib "MasterRecord.dll" () As String
Public Declare Function my_sysinformation Lib "temp_inform" () As SysInform_records
Public Declare Function WHDFingerPrint Lib "My_SysInformlib1.dll" () As HDFingerPrint
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
'Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Long, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'********************************
Type condition_type
ty As Integer 'Çø·Öµã(0)xian(1£©Ô²£¨2£©
no As Integer 'µãÏßÔ²µÄÐòºÅ
old_no As Integer
End Type
Global condition_type0 As condition_type
Type condition_data_type
 condition_no As Integer
  condition(1 To 8) As condition_type
   level As Byte
End Type
'***************************************
Type son_data
 last_son As Integer
 son(1 To 16) As condition_type
End Type
Global son_data0 As son_data

Type ratio_for_measure_type
 is_fixed_ratio As Boolean
  Ratio_for_measure As Single
   ratio_for_measure0 As Single '¼ÇÂ¼Ô­Ê¼Öµ
   sons As son_data
End Type
Global Ratio_for_measure As ratio_for_measure_type
Global is_set_Hscroll2_data As Boolean
Type inform_condition_data_type
data As condition_type
wenti_no As Integer
End Type
Global inform_condition_data As inform_condition_data_type
Type depend_element_type
depend_degree As Integer
inter_ty As Integer
element(2) As condition_type
End Type
Global inform_data_base() As condition_type
Global inform_data_last_item As Integer
Global c_data_for_reduce As condition_data_type
Type condition_data0_type
data As condition_data_type
condition_tree_no(1 To 8) As Integer
End Type
Type condition_tree_type
 conditions As condition_data0_type
  condition As condition_type
   conclusion_no As Byte
    depend_dispaly_no(8) As Integer
     father_condition_no(1) As Integer
      temp_father(1) As Integer
       pre_no As Integer
        next_no As Integer
End Type
Global condition_tree() As condition_tree_type
Global condition_tree_head As Integer
Global condition_tree_tail As Integer
Global last_condition_tree  As Integer
Type pre_add_condition_type0
type As Byte
conditions As condition_data_type
End Type
Global pre_add_condition() As pre_add_condition_type0
Type index_type
 i(8) As Integer
End Type
Global condition_data0 As condition_data_type
Type record_data0_type
 condition_data As condition_data_type
  condition_for_value_ As condition_type
    theorem_no As Integer
End Type
Type record_data1_type
display_type As Byte
      index  As index_type
       is_removed As Boolean
         is_proved As Byte
End Type
Type record_data_type
display_string As String
data0 As record_data0_type
data1 As record_data1_type
End Type
'********************
Type record_type
no_reduce As Byte '1 done comnbine,2, ²»ÔÙÍÆÀí
conclusion_no As Byte
conclusion_ty As Byte '=0'=1 ¶¨Öµ,=2¼«Ð¡,=3¼«´ó
display_times As Integer
is_same_th As Boolean
display_no As Integer
depend_display_no(8) As Integer
display_no_ As Integer
branch As Integer
is_depend As Byte '0 ÎÞ¹Ø
index  As index_type ' ±£´æÊý¾Ý
End Type
Type total_record_type
record_data As record_data_type
record_ As record_type
End Type
'***********
Type record_type0
 condition_data As condition_data_type
    to_no(1) As Integer
     para(1) As String
End Type
Global record_00 As record_type0
Global record_data0 As record_data_type
Global record0 As total_record_type
Global record_0 As record_data_type
Global record1 As record_data_type
Type note_space_type
top_left As POINTAPI
bottem_right As POINTAPI
End Type
Global note_space() As note_space_type
'Global con_Circ(4) As con_circle_data_type
'Global last_con_circle As Integer
Global Polygon_for_change As polygon_for_change_type
Type circle_for_change_type
move As POINTAPI
rote_angle As Single
similar_ratio As Single
c As Integer
c_coord As POINTAPI
poi_coordinate(1 To 10) As POINTAPI
radii As Long
direction As Integer
End Type
Global Circle_for_change As circle_for_change_type
'****************
Type V_line_value_data0_type
 v_line As Integer
  v_poi(1) As Integer
   value As String
    unit_value As String
    line_value_no As Integer
     record As record_data_type
End Type
Type V_line_value_type
data(8) As V_line_value_data0_type
record_ As record_type
End Type
Global V_line_value() As V_line_value_type
Global con_V_line_value(3) As V_line_value_type
Type V_two_line_time_value_data0_type
 poi(3) As Integer
 line_no(1) As Integer
 n(3) As Integer
 value As Integer
 record As record_data_type
End Type
Type V_two_line_time_value_type
data(8) As V_two_line_time_value_data0_type
record_ As record_type
End Type
Global V_two_line_time_value() As V_two_line_time_value_type
Global con_V_two_line_time_value(3) As V_two_line_time_value_type
'****************************************
Type tri_function_data_type
A As Integer
sin_value As String
cos_value As String
tan_value As String
ctan_value As String
initial_data As Byte
record As record_data_type
End Type
Type tri_function_type
 data(8) As tri_function_data_type
  record_ As record_type
End Type
Global tri_function() As tri_function_type
'***********
Type segment_type
 line_no As Integer
 poi(1) As Integer
 n(1) As Integer
 para As String
End Type
Type length_of_polygon_type0
last_segment As Integer
segment(15) As segment_type
value As String
Area As String
record As record_data_type
End Type
Type length_of_polygon_type
polygon_ty As Byte
polygon_no As Integer
data(8) As length_of_polygon_type0
record_ As record_type
End Type
Global length_of_polygon() As length_of_polygon_type
Global con_length_of_polygon(3) As length_of_polygon_type
Type factor0_type
para As String
last_factor As Integer
factor() As String
order() As Integer
End Type
Type factor_type
data(8) As factor0_type
End Type
Type value_string0_type
value As String
factor  As factor_type
index(1) As Integer
End Type
'***********************
Type value_string_type
data(8) As value_string0_type
End Type
Global Dvalue_string() As value_string_type
Type angle_data_type
cond_data As condition_data_type
other_no As Integer
poi(2) As Integer
line_no(1) As Integer
te(1) As Byte
total_no As Integer
total_no_ As Byte 'È«½Ç±àºÅ
index(1) As Integer
no_reduce As Byte
value As String
value_no As Integer 'angle3_value µÄºÅ
End Type
Global angle_data_0 As angle_data_type
Type angle_type
data(8) As angle_data_type
End Type
'******************
Type angle_no_type
 no As Integer
 sh As Boolean 'is sharp angle
End Type
Type total_angle_data_type
line_no(1) As Integer
inter_point As Integer
angle_no(3) As angle_no_type 'Ò»¸öÈ«½Ç¶ÔÓ¦4¸öÆÕÍ¨½Ç
is_used_no As Integer
index(1) As Integer
record_for_value As record_data0_type
value As String
value_no As Integer 'angle3_value µÄºÅ
is_draw_verti_mark As Boolean
End Type
Type total_angle_type
data(8) As total_angle_data_type
End Type
'**************************
Global angle_data0 As angle_data_type
Type add_condition_type
condition_type As Byte
condition_no As Integer
End Type
Global add_condition() As add_condition_type
Global last_add_conditions As Integer
'¼ÇÂ¼Ö¤Ã÷¹ý³ÌÖÐÔö¼ÓµÄÌõ¼þ
Type same_three_line_data_type
line_no(2) As Integer
record As record_data_type
End Type
Type same_three_lines_type
data(8) As same_three_line_data_type
record_ As record_type
End Type
Global same_three_line_data As same_three_line_data_type
Global same_three_lines() As same_three_lines_type
'**************************
Global angle() As angle_type
Global last_condition_angle As Integer
Global T_angle() As total_angle_type
'***************
Type new_point_data_type
 poi(1) As Integer
  add_to_line(1) As Integer
   add_to_circle(1) As Integer
    record As record_data_type
     cond As condition_type 'Ìí¼Ó¸¨ÖúµãÒý½øµÄÊý¾Ý
       display_string As String
End Type
Global new_point_data_0 As new_point_data_type
Type new_point_type
 data(8) As new_point_data_type
     record_ As record_type
End Type
'Global new_point_data_0 As Integer
Global new_point() As new_point_type
'*********
Type relation_from_line_to_triangle_data_type
 poi(3) As Integer
  n(3) As Integer
   line_no(1) As Integer
    triangle(1) As Integer
     record As record_data_type
End Type
Type relation_from_line_to_triangle_type
data(8) As relation_from_line_to_triangle_data_type
record_ As record_type
End Type
Global relation_from_line_to_triangle() As relation_from_line_to_triangle_type
 '************
Type relation_from_triangle_to_line_type
data(8) As relation_from_line_to_triangle_data_type
record_ As record_type
End Type
Global relation_from_triangle_to_line() As relation_from_triangle_to_line_type
'Global last_relation_from_triangle_to_line(1) As Integer
'Global old_last_relation_from_triangle_to_line As Integer
'Global last_relation_from_triangle_to_line_for_aid(1) As Integer
 
'**********************
Type Av_no_type
 no As Integer
  index As Integer
End Type
Type temp_no_type
av_no() As Av_no_type
End Type
Global angle_value As temp_no_type
Global angle_value_90 As temp_no_type
Global angle_value_150 As temp_no_type
Global three_angle_value_sum As temp_no_type
Global two_angle_value_90 As temp_no_type
Global two_angle_value_180 As temp_no_type
Global two_angle_value_sum As temp_no_type
'*******
Type point_pair_data0_type
poi(11) As Integer '8-11ºÏ±ÈÀàÐÍ
n(11) As Integer
old_n(11) As Integer
con_line_type(1) As Byte
line_no(5) As Integer
is_h_ratio As Byte
End Type
Type point_pair_data_type
data0 As point_pair_data0_type
record As record_data_type
End Type
'***********
Global dpoint_pair_data_0 As point_pair_data_type
Global dp_data0 As point_pair_data0_type
 Type Dpoint_pair_type
  reduce As Boolean
  data(8) As point_pair_data_type 'poi(11) As Integer '8-11ºÏ±ÈÀàÐÍ
   record_ As record_type
    similar_triangle_no As Integer
     'posible_similar_no As Integer
      'Èç¹ûÊÇÁ½¸öÏß¶Î±È¹¹³É£¬²»ÔÙºóÍÆtrue
     reduce_type As Boolean
 End Type
 Type con_Dpoint_pair_type
  data(8) As point_pair_data0_type 'poi(11) As Integer '8-11ºÏ±ÈÀàÐÍ
 End Type
 Global con_dpoint_pair(3) As con_Dpoint_pair_type
 Global Ddpoint_pair() As Dpoint_pair_type
 Global pseudo_dpoint_pair() As Dpoint_pair_type
'*******************
Type two_order_equation_data_type
para(2) As String
roots(1) As String
record As record_data_type
End Type
Global two_order_equation_data_0 As two_order_equation_data_type
Type two_order_equation_type
data(8) As two_order_equation_data_type
record_ As record_type
End Type
Global two_order_equation() As two_order_equation_type
Global con_two_order_equation(3) As two_order_equation_type
Global Deangle As temp_no_type
'mid_point_line 7
Type mid_point_line_data_type
 poi(5) As Integer
 record As record_data_type
End Type
Global mid_point_line_data_0 As mid_point_line_data_type
Type mid_point_line_type
 data(8) As mid_point_line_data_type
  record_ As record_type
End Type
Global mid_point_line() As mid_point_line_type
Type con_mid_point_line_type
 data(8) As mid_point_line_data_type
End Type
Global con_mid_point_line(3) As con_mid_point_line_type
'*********
Type Equation_data0_type
para_xx As String
para_yy As String
para_xy As String
para_x  As String
para_y As String
para_c As String
root(1) As String
record As record_data_type
End Type
Type Equation_type
data(8) As Equation_data0_type
record_ As record_type
End Type
Global equation() As Equation_type
'*********
Type Equation_group_data_type
equation(2) As Integer
para_xx As String
root(0 To 3, 0 To 1) As String
record As record_data_type
End Type
Type Equation_group_type
data(8) As Equation_group_data_type
record_ As record_type
End Type
Global equation_group() As Equation_group_type
Type relation_string_type0
relation_string As String
element(2) As String
para(2) As String
record As record_data_type
c_data As condition_data_type
End Type
Type relation_string_type
data(8) As relation_string_type0
record_ As record_type
End Type
Global relation_string() As relation_string_type
'*********
Type eline_data0_type
 poi(3) As Integer
 line_no(1) As Integer
 n(3) As Integer
 eside_tri_point(2) As Integer
End Type
Type eline_data_type
data0 As eline_data0_type
record As record_data_type
End Type
Global eline_data_0 As eline_data_type
Type eline_type
 data(8) As eline_data_type
  reduce As Boolean 'true set_reduce false
  record_ As record_type
   triangle As Integer
    direction As Integer
 End Type
 Global eline_data0 As eline_data0_type
 Global Deline() As eline_type
 Global pseudo_eline() As eline_type
 Global con_eline(3) As eline_type
 Global last_condition_eline As Integer
'four_point_on_circle '9
Type three_point_on_circle_data_type
circ As Integer
 poi(2) As Integer
  record As record_data_type
End Type
Type three_point_on_circle_type
data(8) As three_point_on_circle_data_type
record_ As record_type
End Type
Global three_point_on_circle() As three_point_on_circle_type
'****************
Type four_point_on_circle_data_type
circ As Integer
 poi(3) As Integer
  angle(3) As Integer
   angle_pair(3, 1) As Integer
   lin_value_no(3) As Integer
    record As record_data_type
End Type
Global p4_on_C As four_point_on_circle_data_type
Type four_point_on_circle_type
data(8) As four_point_on_circle_data_type
record_ As record_type
End Type
 Type con_four_point_on_circle_type
  data(8) As four_point_on_circle_data_type
 End Type
Global con_Four_point_on_circle(3) As con_four_point_on_circle_type
Global four_point_on_circle() As four_point_on_circle_type
'**************
Global angle_relation As temp_no_type
Type in_verti_type
line_no As Integer
inter_point As Integer
verti_no As Integer
End Type
Type in_paral_type
line_no As Integer
paral_no As Integer
End Type
Type branch_type
branch_no As Integer
can_not_cut As Boolean
End Type
Type tanget_circle_for_line0
 data_no As condition_type
  tangent_circ As Integer
   tangent_poi As Integer
End Type
Type tanget_circle_for_line
 tangent_circ_no As Integer
  tangent_circle(1 To 8) As tanget_circle_for_line0
End Type
'****************************
Type tangent_element_data0
tangent_point As Integer
tangent_element_no As Integer
End Type
Type tangent_element_data
element_no As Integer
tangent_celement(1 To 4) As tangent_element_data0
End Type
'***************************************************
Type Diameter_type
poi(1) As Integer
cond As condition_data_type
End Type
Type parent_data_type
last_element As Byte
element(1 To 3) As condition_type
related_circle As Integer
related_point(2) As Integer
co_degree As Byte
inter_type As Integer 'Èç¹ûÊÇ½»µã£¨Ïß£¬Ïß£¬Ô²£©£¬¼ÇÂ¼ÆäÀàÐÍ
ratio As Single
End Type

'****************************************************
Type circle_data0_type
 center As Integer
  c_coord As POINTAPI
   real_radii As Single 'ÓÐ¶¨°ë¾¶µÄÔ²
    radii As Long '
     color As Byte
      name As String * 2
       visible As Byte
        in_point(10) As Integer
          tangent_circle As tangent_element_data
           tangent_line As tangent_element_data
End Type
'***********************************
Type circle_data_type
data0 As circle_data0_type
parent As parent_data_type
sons As son_data
from_wenti_no As Integer
depend_element As depend_element_type
depend_poi(8) As Integer
degree As Byte
Diameter(1 To 4) As Diameter_type
last_Diameter As Integer
radii_no As Integer
radii_depend_poi(1) As Integer '=-1 ¹Ì¶¨µÄ°ë¾¶
depend_para As Single
inform As String
other_no As Integer
circle_para(4) As String
circ_coord(1) As String
circ_radii As String
circ_radii_record As condition_data0_type
circ_radii_squre As String
circ_radii_squre_record As condition_data0_type
is_change As Boolean
input_type As Byte
circle_type As Byte
End Type
'******************************
Type circle_type
is_set_data As Boolean
data(8) As circle_data_type
End Type
'**********************************
Type temp_circle_for_input
is_using As Boolean
data(1) As circle_data0_type
End Type
Global m_temp_circle_for_input As temp_circle_for_input
Global circ_0 As circle_type
Type con_circle_data_type
 center As Integer
  c_coord As POINTAPI
   radii As Long
    color As Byte
     name As String * 2
      visible As Byte
       circle_no As Integer
'       in_point(10) As Integer
End Type
Type io_circle_data_type
circle_no As Integer
input_ty As Byte
circle_ty As Byte
circle_data As circle_data0_type
End Type
Type line_data0_type
 poi(1) As Integer 'È·¶¨Ö±ÏßµÄ×îÏÈÁ½¸öµã
  depend_poi(1) As Integer
   depend_poi1_coord As POINTAPI 'µ¥µã¾ö¶¨Ö±ÏßµÄ£¬µÚ¶þµã×ø±ê
  end_point_coord(1) As POINTAPI '¶Ëµã×ø±ê
  in_point(10) As Integer 'Ö±ÏßÉÏÒÀÐòÅÅÁÐµÄµã
   color(1 To 9) As Byte '
    total_color As Byte
     visible As Byte '
      type As Byte ' Ìõ¼þ»ò½áÂÛ
       tangent_circle As tangent_element_data
End Type
'***************************************

'***************************************
'¼ÇÂ¼Ö±ÏßÊý¾Ý,¹©ÍÆÀíÓÃ
Type line_data_type
data0 As line_data0_type
parent As parent_data_type
sons As son_data
condition As Byte
k(1) As Single
eangle_no As Integer
tangent_line_no As Integer 'Èç¹ûÊÇÇÐÏß£¬¼ÇÂ¼ÇÐÏßÐòºÅ
in_paral(8) As in_paral_type
in_verti(8) As in_verti_type
inform As String
line_para(2) As String 'Ò»°ãÊ½Ö±Ïß·½³Ì
is_change As Boolean
other_no As Integer 'con_line
cond_data As condition_data_type
End Type
'************************************************
'Global line_data_0 As line_data_type
'******************************************
Type line_type
is_set_data As Boolean
data(8) As line_data_type
End Type
Type io_line_data_type
line_no As Integer
depend_element As depend_element_type
condition As Byte
line_data As line_data0_type
End Type
Type temp_line_for_input
 is_using As Boolean
 data(1) As line_data0_type
End Type
Global m_temp_line_for_input As temp_line_for_input
Global last_condition_line As Integer
'mid_point 12
Type mid_point_data0_type
poi(2) As Integer
n(2) As Integer
line_no As Integer
End Type
Type mid_point_data_type
data0 As mid_point_data0_type
record As record_data_type
End Type
Global mid_point_data_0 As mid_point_data_type
Type mid_point_type
data(8) As mid_point_data_type
reduce As Boolean
record_ As record_type
End Type
Type con_mid_point_type
data(8) As mid_point_data0_type
 End Type
Global Dmid_point_data0 As mid_point_data0_type
Global Dmid_point() As mid_point_type
Global con_mid_point(3) As con_mid_point_type
Global pseudo_mid_point() As mid_point_type
 'para 13
Type two_line_type
line_no(1) As Integer
inter_poi As Integer
record As record_data_type
End Type
Global two_line_data_0 As two_line_type
Type paral_data0_type
data0 As two_line_type
distance As String ' Æ½ÐÐÏß¼äµÄ¾àÀë
distance_no As Integer
End Type
Type paral_type
reduce As Boolean
data(8) As paral_data0_type
'old_data As two_line_type
record_ As record_type
End Type
Global Dparal() As paral_type
Type con_paral_type
data(8) As two_line_type
 End Type
Global con_paral(3) As con_paral_type
'parallelogram 14
'Type equal_side_triangle_data0_type
'poi(2) As Integer
'triangle As Integer
'dir As Integer
'End Type
'Type equal_side_triangle_data_type
'data0 As equal_side_triangle_data0_type
'record_ As record_data_type
'triangle As Integer
'dir As Integer
'End Type
'Type equal_side_triangle_type
'data(8) As equal_side_triangle_type
'record_ As record_type
'End Type
'Global Dequal_side_triangle() As equal_side_triangle_type
'Type con_equal_side_triangle_type
'data(8) As equal_side_triangle_data_type
'End Type
'Global con_equal_side_triangle(4) As con_equal_side_triangle_type
Type polygon4_data_type
ty As Byte
no As Integer
poi(4) As Integer '×îºóÒ»µã
line_no(3) As Integer
dia_poi(3) As Integer '¶Ô½ÇÏß
angle(3) As Integer
midpoi(3) As Integer
midpoi_no(3) As Integer
start_poi As Byte ' ÓÃÓÚÌÝÐÎ
triAngle1(1) As Integer
triAngle2(1) As Integer
area_value As String '
area_value_ As String '¼ÇÂ¼²»º¬Î´ÖªÊýµÄ½á¹û
area_value_no As Integer
condition As condition_type
tangent_circle As Integer
tangnet_poi(3) As Integer
height_no As Integer
two_buttom_sun(1) As condition_type
index  As Integer
End Type
Type polygon4_type
data(8) As polygon4_data_type
End Type
Global Dpolygon4() As polygon4_type
Type dpolygon4_type
polygon4_no As Integer
record As record_data_type
End Type
Global dpolygon4_data_0 As dpolygon4_type
Type parallelogram_type
data(8) As dpolygon4_type
record_ As record_type
 End Type
Global Dparallelogram() As parallelogram_type
Type con_parallelogram_type
data(8) As polygon4_data_type
End Type
Global con_parallelogram(3) As con_parallelogram_type
'Global last_condition_parallelogram As Integer
Global polygon4_data0 As polygon4_data_type
'point
'Global old_last_conditions.last_cond(1).point_no As Integer
'Global last_conditions.last_cond(1).point_no_for_aid As Integer
'Global old_last_conditions.last_cond(1).point_no_for_aid As Integer
Type relation_on_line_data0_type
poi(3) As Integer
value As String
line_no  As Integer
End Type
Type relation_on_line_data_type
data0 As relation_on_line_data0_type
record  As condition_data_type
End Type
Global relation_on_line_data_0 As relation_on_line_data_type
Type relation_on_line_type
data(8) As relation_on_line_data_type
End Type
Global relation_on_line_data0 As relation_on_line_data0_type
Global relation_on_line() As relation_on_line_type

'relation  '16
Type relation_data0_type
poi(5) As Integer
n(5) As Integer
value As String
value_ As String
line_no(2) As Integer
ty As Byte
End Type
Type relation_data_type
data0 As relation_data0_type
record As record_data_type
End Type
Global relation_data_0 As relation_data_type
Type relation_type
data(8) As relation_data_type
reduce As Boolean
record_ As record_type
End Type
Global relation_data0 As relation_data0_type
Global Drelation() As relation_type
Type con_relation_type
data(1) As relation_data0_type
 End Type
Global con_relation(3) As con_relation_type '½áÂÛ
'Global last_condition_relation As Integer
'******************************
Type V_relation_data0_type
v_line(2) As Integer
line_no(2) As Integer
value As String
value_ As String
ty As Byte
End Type
Type V_relation_data_type
data0 As V_relation_data0_type
record As record_data_type
End Type
Type v_relation_type
data(8) As V_relation_data_type
record_ As record_type
End Type
Global v_Drelation() As v_relation_type
'similar_triangle '17
Global pseudo_relation() As relation_type
Type two_triangle_type
 triangle(1) As Integer
  direction As Integer
   record As record_data_type
End Type
Type similar_triangle_type
 data(8) As two_triangle_type
  record_ As record_type
End Type
Global two_triangle0 As two_triangle_type
Global Dsimilar_triangle() As similar_triangle_type
Type con_similar_triangle_type
 data(8) As two_triangle_type
End Type
Global con_similar_triangle(3) As con_similar_triangle_type
Global last_condition_similar_triangle As Integer
Type total_equal_triangle_type
record_ As record_type
data(8) As two_triangle_type
End Type

Type con_total_equal_triangle_type
data(8) As two_triangle_type
End Type
Global Dtotal_equal_triangle() As total_equal_triangle_type
Global con_total_equal_triangle(3) As con_total_equal_triangle_type
'***************************Global Last_total_equal_triangle(1) As Integer
Type pseudo_two_triangle_data_type
'ÄâÈ«µÈ
triA(1) As temp_triangle_data_type 'Èý½ÇÐÎ¶Ô
pseudo_point(1) As Integer 'ÄâÈ«µÈµÄÄâ¶ÔÓ¦µã
conclusion_poi(1) As Integer
record As record_data_type '¼ÇÂ¼
pseudo_condition_data  As condition_data_type '(0),(1) pseudo condition'ÄâÈ«µÈÌõ¼þ
ty As Byte 'ÀàÐÍ
two_triA_n As Integer
e_A_n As Integer
e_l_n(2) As Integer
e_l_ty As Byte
is_set As Byte
End Type
Type pseudo_two_triangle_type
data(8) As pseudo_two_triangle_data_type
record_ As record_type
End Type
Global pseudo_total_equal_triangle() As pseudo_two_triangle_type
Global pseudo_similar_triangle() As pseudo_two_triangle_type
'**************************** area_relation 4
Type area_relation_data_type 'Ãæ»ý±È
 area_element(2) As condition_type
  value As String
record As record_data_type
End Type
Global area_relation_data_0 As area_relation_data_type
    Type area_relation_type
data(8) As area_relation_data_type
record_ As record_type
End Type
Global Darea_relation() As area_relation_type
'Global Last_condition_area_relation As Integer
Global con_area_relation(3) As area_relation_type
Type equal_or_simlilar_triangle_data0_type
no As Integer
triangle_no As Integer
direction As Integer
End Type
Type equal_or_simlilar_triangle_data_type
last_E_S_triangle As Integer
E_S_triangle() As equal_or_simlilar_triangle_data0_type
End Type
Type value_from_element
para As String
element_no As Integer
add_value As String
End Type
Type triangle_data0_type
poi(3) As Integer 'poi(3) ×îºóÒ»µã
line_no(2) As Integer
angle(2) As Integer
angle_value(2) As Integer
angle_value_no_from_two_angle(2) As Integer
angle_value_from_two_angle(2)  As value_from_element
tri_function(2) As Integer
line_value(2) As Integer
midpoint_no(2) As Integer
relation_no(2, 1) As condition_type
verti_no(2) As Integer
two_angle_equal(2) As Integer
eangle_no(2, 1) As condition_type
verti_line(2) As Integer
eangle_line(2) As Integer
mid_point_line(2) As Integer
inner_center As Integer
verti_center As Integer
center As Integer
length_of_sides_value As String
length_of_sides  As condition_data_type
time_of_two_line(2) As Integer
re_value(2) As String
height_no(2) As Integer
v_line(2) As Integer
tanget_circle As Integer
tangent_poi(2) As Integer
sum_of_two_sq_line(2) As Byte '±íÊ¾ÒÑ×÷¹´¹É¶¨Àí
right_angle_no As Integer
Area As String
area_no As Integer
index  As index_type
total_equal_triangle As equal_or_simlilar_triangle_data_type
similar_triangle  As equal_or_simlilar_triangle_data_type
condition As condition_type
no_reduce As Byte
End Type
'triangle  '19
Type triangle_type
index As index_type
data(8) As triangle_data0_type
from_ty As Byte
from_no As Integer
ty As Byte
ty_no As Integer 'ÀàÐÍ
epolygon_no As Integer
record_ As record_type '
End Type
Global triangle() As triangle_type
Global triangle_data0 As triangle_data0_type
'Global last_triangle As Integer
Global last_condition_triangle As Integer
Global Area_element_in_conclusion() As condition_type '¼ÇÂ¼ÍÆÀí¹ý³Ì²úÉúµÄÃæ»ý
Global last_area_element_in_conclusion  As Integer
Type Rtriangle_data_type
triangle As Integer
direction As Integer
record As record_data_type
End Type
Type Rtriangle_type
data(8) As Rtriangle_data_type
record_ As record_type
End Type
Global Rtriangle() As Rtriangle_type
'Global last_Rtriangle As Integer
'Global last_Rtriangle_for_aid As Integer
'three_point_on_line  '20
Type three_point_on_line_data_type
record As record_data_type
poi(2) As Integer
is_no_initial  As Byte
End Type
Global three_point_on_line_data_0 As three_point_on_line_data_type
Type three_point_on_line_type
data(8) As three_point_on_line_data_type
record_ As record_type
End Type
Global three_point_on_line() As three_point_on_line_type
Type con_three_point_on_line_type
data(8) As three_point_on_line_data_type
 End Type
Global con_Three_point_on_line(3) As con_three_point_on_line_type
'Global last_condition_Three_point_on_line As Integer
Type two_point_conset_data_type
 poi(1) As Integer
 record As record_data_type
End Type
Type two_point_conset_type
 data(8) As two_point_conset_data_type
 record_ As record_type
End Type
Global two_point_conset() As two_point_conset_type
'two_line_value  '21
Type two_line_value_data0_type
 poi(5) As Integer
 n(5) As Integer
  para(1) As String
   line_no(2) As Integer
    value As String
     value_ As String
End Type

Type two_line_value_data_type
data0 As two_line_value_data0_type
record As record_data_type
End Type
Type two_line_value_type
data(8) As two_line_value_data_type
reduce As Boolean
record_ As record_type
End Type
Global two_line_value_data0 As two_line_value_data0_type
Global two_line_value() As two_line_value_type
Type con_two_line_value_type
data(8) As two_line_value_data0_type
End Type
Global con_two_line_value(3) As con_two_line_value_type
'**********************************************
'Global last_conditions.last_cond(1).two_line_value_no As Integer
'Global last_condition_two_line_value As Integer
'Global old_last_two_line_value As Integer
'Global last_two_line_value_for_aid(1) As Integer
'Global temp_last_two_line_value As Integer

'line3_value'22
Type line3_value_data0_type
 poi(9) As Integer
  n(9) As Integer
    para(2) As String
     line_no(4) As Integer
       value As String
        value_ As String
End Type
Type line3_value_data_type
data0 As line3_value_data0_type
record As record_data_type
End Type
Type line3_value_type
data(8) As line3_value_data_type
reduce As Boolean
 record_ As record_type
End Type
Global line3_value_data0 As line3_value_data0_type
Global line3_value() As line3_value_type
Type con_line3_value_type
data(8) As line3_value_data0_type
End Type
Global con_line3_value(3) As con_line3_value_type
Global pseudo_line3_value() As line3_value_type
'Global last_conditions.last_cond(1).line_no3_value(1) As Integer
'Global old_last_conditions.last_cond(1).line_no3_value As Integer
'Global last_conditions.last_cond(1).line_no3_value_for_aid(1) As Integer
'Global temp_last_conditions.last_cond(1).line_no3_value As Integer
'Global last_condition_line3_value As Integer
Global Two_angle_value As temp_no_type
Global three_angle_value As temp_no_type
'Global two_angle_value_sum As temp_no_type
Global three_angle_value0 As temp_no_type
Global Two_angle_value0 As temp_no_type
'Global Two_angle_value_90 As temp_no_type
'Global Two_angle_value_180 As temp_no_type

'last_verti  '24
Type verti_type
reduce As Boolean
data(8) As two_line_type
'old_data As two_line_type
record_ As record_type
End Type
Global Dverti() As verti_type
Type con_verti_type
data(8) As two_line_type
End Type
Global con_verti(3) As con_verti_type
'Global last_conditions.last_cond(1).verti_no As Integer
'Global old_last_verti As Integer
'Global last_verti_for_aid(1) As Integer
'Global temp_last_verti As Integer
'Global last_condition_verti As Integer
Type arc_data_type
 poi(1) As Integer
  cir As Integer
   small_or_big As Boolean
    index(1) As Integer
     old_index(1) As Integer
End Type
Global arc_data_0 As arc_data_type
Type arc_type
data(8) As arc_data_type
End Type

Global arc() As arc_type
'arc_value  '25
Type arc_value_data_type
arc As Integer
'   cir As Integer
    value As String
 record As record_data_type
End Type
Global arc_value_data_0 As arc_value_data_type
Type arc_value_type
data(8) As arc_value_data_type
 record_ As record_type
End Type
Global arc_value() As arc_value_type
Type con_arc_value_type
data(8) As arc_value_data_type
End Type
Global con_arc_value(3) As con_arc_value_type
'Global last_arc_value As Integer
'Global old_last_arc_value As Integer
'Global last_arc_value_for_aid As Integer
'Global temp_last_arc_value As Integer
'Global last_condition_arc_value As Integer

'equal_arc  '26
Type equal_arc_data_type
 arc(2) As Integer
  record As record_data_type
End Type
Global equal_arc_data_0 As equal_arc_data_type
Type equal_arc_type
 data(8) As equal_arc_data_type
 record_ As record_type
End Type
Type con_equal_arc_type
 data(8) As equal_arc_data_type
End Type
Global equal_arc() As equal_arc_type
Global con_equal_arc(3) As con_equal_arc_type
'Global last_condition_equal_arc As Integer
'ratio_of_two_arc '27
Type ratio_of_two_arc_data_type
 poi(3) As Integer
  circle As Integer
   value As String
 record As record_data_type
End Type
Type ratio_of_two_arc_type
data(8) As ratio_of_two_arc_data_type
 record_ As record_type
End Type
Global ratio_of_two_arc() As ratio_of_two_arc_type
Type con_ratio_of_two_arc_type
data(8) As ratio_of_two_arc_data_type
End Type
Global con_ratio_of_two_arc(3) As con_ratio_of_two_arc_type
'Global last_condition_ratio_of_two_arc As Integer
'angle_less_angle  '28
Type angle_less_angle_data_type
angle(1) As Integer
record As record_data_type
End Type
Type angle_less_angle_type
data(8) As angle_less_angle_data_type
record_  As record_type
End Type
Global angle_less_angle() As angle_less_angle_type
Type con_angle_less_angle_type
data(8) As angle_less_angle_data_type
End Type
Global con_angle_less_angle(3) As con_angle_less_angle_type
'Global last_angle_less_angle As Integer
'Global old_last_angle_less_angle As Integer
'Global last_angle_less_angle_for_aid As Integer
'Global temp_last_angle_less_angle As Integer
'Global last_condition_angle_less_angle As Integer
'line_less_line =  '29
Type line_less_line_data_type
poi(3) As Integer
record As record_type
record_0 As record_type
End Type
Type line_less_line_type
data(8) As line_less_line_data_type
record_  As record_type
End Type
Global line_less_line() As line_less_line_type
Type con_line_less_line_type
data(8) As line_less_line_data_type
End Type
Global con_line_less_line(3) As line_less_line_type
'Global last_conditions.last_cond(1).line_no_less_line As Integer
'Global old_last_conditions.last_cond(1).line_no_less_line As Integer
'Global last_conditions.last_cond(1).line_no_less_line_for_aid As Integer
'Global temp_last_conditions.last_cond(1).line_no_less_line As Integer
'Global last_condition_line_less_line As Integer

'line_less_line2  '30
Type line_less_line2_data_type
poi(5) As Integer
record As record_data_type
'record_0 As record_type
End Type
Type line_less_line2_type
data(8) As line_less_line2_data_type
record_ As record_type
End Type
Global line_less_line2() As line_less_line2_type
Type con_line_less_line2_type
data(8) As line_less_line2_data_type
End Type
Global con_line_less_line2(3) As con_line_less_line2_type
'Global last_conditions.last_cond(1).line_no_less_line2 As Integer
'Global old_last_conditions.last_cond(1).line_no_less_line2 As Integer
'Global last_conditions.last_cond(1).line_no_less_line2_for_aid As Integer
'Global temp_last_conditions.last_cond(1).line_no_less_line2 As Integer
'Global last_condition_line_less_line2 As Integer
'line2_less_line2  '31
Type line2_less_line2_data_type
poi(7) As Integer
record As record_data_type
End Type
Type line2_less_line2_type
data(8) As line2_less_line2_data_type
record_ As record_type
End Type
Global line2_less_line2() As line2_less_line2_type
Type con_line2_less_line2_type
data(8) As line2_less_line2_data_type
End Type
Global con_line2_less_line2(3) As con_line2_less_line2_type
'Global last_conditions.last_cond(1).line_no2_less_line2 As Integer
'Global old_last_conditions.last_cond(1).line_no2_less_line2 As Integer
'Global last_conditions.last_cond(1).line_no2_less_line2_for_aid As Integer
'Global temp_last_conditions.last_cond(1).line_no2_less_line2 As Integer
'Global last_condition_line2_less_line2 As Integer
Type equal_3angle_type0
angle(3) As Integer
line_no(3) As Integer
poi As Integer
record As record_data_type
End Type
Type equal_3angle_type
data(8) As equal_3angle_type0
End Type
Global equal_3angle() As equal_3angle_type
'angle3_value  '32
Type angle3_value_data0_type
 angle(5) As Integer 'angle(3) ¼ÇÂ¼Ç°Á½½ÇºÍ
  angle_(3 To 5) As Integer
  value As String
  value_ As String
    para(2) As String
      ty(2) As Byte '¼ÇÂ¼Á½½Ç µÄ¹ØÏµ
       ty_(2) As Byte
       type As Byte '¼ÇÂ¼ÀàÐÍ ÏàµÈ,..
        total_angle(2) As Integer '
         total_angle3_value_no As Integer
          reduce  As Boolean ' ,¹©ÍÆÀí
           no_zero_angle As Byte
            no_combine As Boolean '²»×öÁ½½ÇºÏ²¢
End Type
Type angle3_value_data_type
data0 As angle3_value_data0_type
           record As record_data_type
End Type
Global angle3_value_data_0 As angle3_value_data_type
Type angle3_value_type
data(8) As angle3_value_data_type
'old_data As angle3_value_data_type
reduce As Boolean
record_ As record_type
 End Type
 ''''''''
 Global temp_angle3_value(15) As angle3_value_data0_type
Global last_temp_a3_v_no As Integer
'
Global con_angle3_value(3) As angle3_value_type
Global angle3_value_data0 As angle3_value_data0_type
Global angle3_value() As angle3_value_type
'*****************
'*******************
'line_value  '33
Type line_value_data0_type
 poi(1) As Integer
  n(1) As Integer
   line_no As Integer
    value As String
     value_ As String
       squar_value As String
End Type
Type line_value_data_type
data0 As line_value_data0_type
record As record_data_type
End Type
Type line_value_type
data(8) As line_value_data_type
reduce As Boolean 'ÊÇ·ñÖ±½ÓÊäÈë
record_ As record_type
End Type
Global line_value_data0 As line_value_data0_type
Global line_value() As line_value_type
Global con_line_value(3) As line_value_type
'Global temp_last_conditions.last_cond(1).line_no_value As Integer
'Global last_condition_line_value As Integer
'tangent_line  '34
Type tangent_line_data_type
poi(1) As Integer
coordinate(1) As POINTAPI
old_coordinate(1) As POINTAPI
new_coordinate(1) As POINTAPI
n(1) As Integer
circ(1) As Integer
ele(1) As condition_type
record As record_data_type
line_no As Integer
visible As Byte '0²»ÏÔÊ¾,1¿ÉÏÔÊ¾,2ÏÔÊ¾,3Î´ÈëÊý¾Ý¿â£¬µÈ´ýÏÔÊ¾,4µ¥¶ËµãÇÐÏß,5
End Type
Type tangent_line_type
data(8) As tangent_line_data_type
tangent_type  As Integer 'ÇÐÏßÀàÐÍ
is_display_in_wenti_data As Boolean 'ÊÇ·ñÒÑÔÚÊäÈëÓï¾äÖÐÏÔÊ¾
'old_data As tangent_line_data_type
in_out_tangent As Byte
record_ As record_type
End Type
Global tangent_line_data0 As tangent_line_data_type
Global tangent_line() As tangent_line_type
Type con_tangent_line_type
data(8) As tangent_line_data_type
'old_data As tangent_line_data_type
in_out_tangent As Byte
'record_ As record_type
End Type
Global con_tangent_line(3) As con_tangent_line_type
'Global last_tangent_line(1) As Integer
'Global old_last_tangent_line As Integer
'Global last_tangent_line_for_aid(1) As Integer
'Global temp_last_tangent_line As Integer
'Global last_condition_tangent_line As Integer
Type tangent_circle_data0
tangent_coord(1) As POINTAPI 'ÇÐµã×ø±ê
circle_center  As POINTAPI
circle_radii As Long
visible As Byte '0²»ÏÔÊ¾,1¿ÉÏÔÊ¾,2ÏÔÊ¾,3Î´ÈëÊý¾Ý¿â£¬µÈ´ýÏÔÊ¾,4µ¥¶ËµãÇÐÏß,5
End Type
'equal_area_triangle  '35
Type tangent_circle_data_type
data0(1) As tangent_circle_data0
center  As Integer
ele(1) As condition_type 'ÇÐÓÚÁ½Ô²,»òÁ½Ïß
tangent_poi(1) As Integer 'ÇÐµã
record As record_data_type
circle_no As Integer 'ÇÐÔ²µÄÐòºÅ
tangent_circle_ty As Integer
End Type
Type tangent_circle_type
data(8) As tangent_circle_data_type
record_ As record_type
End Type
Global m_tangent_circle() As tangent_circle_type
Type con_tangent_circle_type
data(8) As tangent_circle_data_type
End Type
Global con_tangent_circle(3) As con_tangent_circle_type
'Global last_tangent_circle As Integer
'Global old_last_tangent_circle As Integer
'Global last_tangent_circle_for_aid As Integer
'Global temp_last_tangent_circle As Integer
'Global last_condition_tangent_circle As Integer
'************
'Type equal_area_triangle_data_type
'triangle(2) As Integer
'record As record_data_type
'End Type
'Global equaL_area_triangle_data_0 As equal_area_triangle_data_type
'Type equal_area_triangle_type
'data(8) As equal_area_triangle_data_type
'record_ As record_type
'End Type
'Global equal_area_triangle() As equal_area_triangle_type
'Type con_equal_area_triangle_type
'data(8) As equal_area_triangle_data_type
'End Type
'Global con_equal_area_triangle(3) As con_equal_area_triangle_type
'Global last_equal_area_triangle As Integer
'Global old_last_equal_area_triangle As Integer
'Global last_equal_area_triangle_for_aid As Integer
'Global temp_last_equal_area_triangle As Integer
'Global last_condition_equal_area_triangle As Integer
'general_string  '36
Type general_string_data_type
item(3) As Integer
para(3) As String
value As String
value_ As String
combine_two_item(1) As Integer
trans_para_for_display As String '¹«Ê½ÍÆµ¼µÄ±ä»»ÏµÊý
trans_para As String '¹«Ê½ÍÆµ¼µÄ±ä»»ÏµÊý
trans_equal_mark As Byte '=0 = ,1¡Ý,2¡Ü,3£¾ ,4£¼
record As record_data_type
End Type
Type general_string_type
data(8) As general_string_data_type
reduce As Boolean
record_ As record_type
display_con_string As String
End Type
'Type con_general_string_type
'data(8) As general_string_data_type
'End Type
Global general_string() As general_string_type
Global con_general_string(3) As general_string_type
Global con_g_s(3, 20) As Integer
Global last_dis_con_gs(3) As Integer
Global last_dis_gs(3) As Integer
Global dis_gs_no%
'Global last_con_general_string As Integer
'Global last_conditions.last_cond(1).general_string_no As Integer
'Global old_last_general_string As Integer
'Global last_general_string_for_aid(1) As Integer
'Global temp_last_general_string As Integer
'Global last_condition_general_string As Integer
'general_angle_string  '37
Type general_angle_string_data_type
angle(3) As Integer
para(3) As String
trans_para(1) As Integer '¹«Ê½ÍÆµ¼µÄ±ä»»ÏµÊý
combine_two_item(1) As Integer
'conclusion_no As Byte
record As record_data_type
'display_times As Byte
End Type
Type general_angle_string_type
data(8) As general_angle_string_data_type
conclusion_no As Byte
record_ As record_type
'display_times As Byte
End Type
Global general_angle_string() As general_angle_string_type
Type con_general_angle_string_type
data(8) As general_angle_string_data_type
End Type
Global con_general_angle_string(3) As con_general_angle_string_type
'Global last_con_general_angle_string As Integer
'Global last_general_angle_string As Integer
'Global old_last_general_angle_string As Integer
'Global last_general_angle_string_for_aid As Integer
'Global temp_last_general_angle_string As Integer
'Global last_condition_general_angle_string As Integer
Type string_value_data_type
s As String
value As String
is_known_value As Boolean
record As record_data_type
'display_times As Byte
End Type
Global string_value_data_0 As string_value_data_type
Type string_value_type
data(8) As string_value_data_type
End Type
'Global last_string_value As Integer
Global con_string_value(3) As string_value_type '½áÂÛ
Global string_value() As string_value_type
'Global old_last_string_value As Integer
'Global temp_last_string_value As Integer
'Global last_string_value_for_aid As Integer
'equal_side_tixing  '38
'Type tixing_data_type
'poi(3) As Integer
'poly4_no As Integer
'record As record_data_type
'End Type
Type tixing_data_type
ty As Byte
area_value_no As Integer
poi(3) As Integer
mid_poi(1) As Integer
mid_point_no(3) As condition_type
paral_no As Integer
poly4_no As Integer
line_value_no(3) As Integer
'distance_paral_no As Integer'¿ÉÓÃÆ½ÐÐÏß¾àÀë±íÊ¾
buttom_(1) As condition_type
sum_of_two_bottom_no As Integer
mid_position_line_value_no As Integer
record As record_data_type
End Type
Global tixing_data_0 As tixing_data_type
Type tixing_type
data(8) As tixing_data_type
'poi(3) As Integer
'midpoi(1) As Integer
'midpoi_no(1) As Integer
'record As record_type
record_  As record_type
End Type
Global Dtixing() As tixing_type
Type con_tixing_type
data(8) As tixing_data_type
'poi(3) As Integer
'midpoi(1) As Integer
'midpoi_no(1) As Integer
'record As record_type
'record_  As record_type
End Type
Global con_Dtixing(3) As con_tixing_type
'Global last_conditions.last_cond(1).tixing_no As Integer
'Global old_last_conditions.last_cond(1).tixing_no As Integer
'Global temp_last_conditions.last_cond(1).tixing_no As Integer
'Global last_conditions.last_cond(1).tixing_no_for_aid As Integer
'Global last_condition_tixing As Integer
'Epolygon 40
'Global equal_side_tixing_data_0 As tixing_data_type
'Type equal_side_tixing_type
'data(8) As tixing_data_type
'record_ As record_type
'End Type
'Type con_equal_side_tixing_type
'data(8) As tixing_data_type
'End Type
'Global con_equal_side_tixing(3) As tixing_type '½áÂÛ
'Global Dequal_side_tixing() As tixing_type

Type epolygon_data_type
p As polygon
no As Integer
circ As Integer
'midpoi(3) As Integer
'midpoi_no(3) As Integer
record As record_data_type
'record_0 As record_type
'display_no As Integer 'Éè±ß³¤µÄÓï¾ä
'display_times As Integer '1 ÐèÒªÏÔÊ¾ 2 ÒÑÏÔÊ¾ 0£¬¸´Ô±
End Type
Global epolygon_data_0 As epolygon_data_type
Type epolygon_type
data(8) As epolygon_data_type
record_ As record_type
'record_0 As record_type
'display_no As Integer 'Éè±ß³¤µÄÓï¾ä
'display_times As Integer '1 ÐèÒªÏÔÊ¾ 2 ÒÑÏÔÊ¾ 0£¬¸´Ô±
End Type

Global epolygon() As epolygon_type
Type con_epolygon_type
data(8) As epolygon_data_type
End Type
Global con_Epolygon(3) As con_epolygon_type
'Global last_conditions.last_cond(1).Epolygon_no As Integer
'Global old_last_conditions.last_cond(1).Epolygon_no As Integer
'Global last_conditions.last_cond(1).Epolygon_no_for_aid As Integer
'Global temp_last_conditions.last_cond(1).Epolygon_no As Integer
'Global last_condition_Epolygon As Integer
'rhombus  '41
'Type rhombus_data_type
'poi(3) As Integer
'midpoi(3) As Integer
'midpoi_no(3) As Integer
'record As record_data_type
'End Type
Type rhombus_type
data(8) As dpolygon4_type
'poi(3) As Integer
'midpoi(3) As Integer
'midpoi_no(3) As Integer
'record As record_type
record_ As record_type
End Type
Global rhombus() As rhombus_type
Type con_rhombus_type
data(8) As dpolygon4_type
'poi(3) As Integer
'midpoi(3) As Integer
'midpoi_no(3) As Integer
'record As record_type
'record_ As record_type
End Type
Global con_rhombus(3) As con_rhombus_type
'Global last_conditions.last_cond(1).rhombus_no As Integer
'Global old_last_conditions.last_cond(1).rhombus_no As Integer
'Global last_conditions.last_cond(1).rhombus_no_for_aid As Integer
'Global temp_last_conditions.last_cond(1).rhombus_no As Integer
'Global last_condition_rhombus As Integer
'long_squre  '42
'Type long_squre_data_type
'poi(3) As Integer
'midpoi(3) As Integer
'midpoi_no(3) As Integer
'record As record_data_type
'record_0 As record_type
'End Type
Type squre0_type
polygon4_no As Integer
four_point_on_circle_no As Integer
length_of_side_no As Integer
length_of_diag_no As Integer
radii_no As Integer
depend_element(1) As depend_element_type
no_reduce As Boolean
is_set_length As Boolean
record As record_data_type
End Type
Type squre_type
data(8) As squre0_type
'poi(3) As Integer
'midpoi(3) As Integer
'midpoi_no(3) As Integer
'record As record_type
record_ As record_type
End Type
Global Dsqure() As squre_type
Type con_squre_type
data(8) As squre0_type
'poi(3) As Integer
'midpoi(3) As Integer
'midpoi_no(3) As Integer
'record As record_type
'record_ As record_type
End Type
Global con_squre(3) As con_squre_type

Type long_squre_type
data(8) As dpolygon4_type
'poi(3) As Integer
'midpoi(3) As Integer
'midpoi_no(3) As Integer
'record As record_type
record_ As record_type
End Type
Global Dlong_squre() As long_squre_type
Type con_long_squre_type
data(8) As dpolygon4_type
'poi(3) As Integer
'midpoi(3) As Integer
'midpoi_no(3) As Integer
'record As record_type
'record_ As record_type
End Type
Global con_long_squre(3) As con_long_squre_type
'Global last_conditions.last_cond(1).last_long_squre_no As Integer
'Global old_last_conditions.last_cond(1).last_long_squre_no As Integer
'Global last_conditions.last_cond(1).last_long_squre_no_for_aid As Integer
'Global temp_last_conditions.last_cond(1).last_long_squre_no As Integer
'Global last_condition_long_squre As Integer
'area_of_triangle  '43
Type area_of_element_data_type
element As condition_type
record As record_data_type
'record_0 As record_type
value As String
value_ As String
End Type
Global area_of_element_data_0 As area_of_element_data_type
Type area_of_element_type
data(8) As area_of_element_data_type
 record_  As record_type
 End Type
Global area_of_element() As area_of_element_type
Type con_area_of_element_type
data(8) As area_of_element_data_type
 'record_  As record_type
 End Type
Global con_Area_of_element(3) As con_area_of_element_type '½áÂÛ
Type two_area_element_value_data_type
 area_element(2) As area_of_element_data_type
 s(1) As String
 value As String
 value_ As String
 record As record_data_type
End Type
Type two_area_element_value_type
data(8) As two_area_element_value_data_type
 record_  As record_type
 End Type
Global two_area_of_element_value() As two_area_element_value_type
'Global last_area_of_triangle As Integer
'Global old_last_area_of_triangle As Integer
'Global last_area_of_triangle_for_aid As Integer
'Global temp_last_area_of_triangle As Integer
'Global last_condition_area_of_triangle As Integer
'area_of_circle  '44
Type area_of_circle_data_type
circ As Integer
record As record_data_type
'record_0 As record_type
value As String
value_ As String
End Type
Global area_of_circle_data_0 As area_of_circle_data_type
Type area_of_circle_type
data(8) As area_of_circle_data_type
 record_  As record_type
 End Type
Global area_of_circle() As area_of_circle_type
Type con_area_of_circle_type
data(8) As area_of_circle_data_type
 'record_  As record_type
 End Type
Global con_Area_of_circle(3) As con_area_of_circle_type '½áÂÛ
'Global last_area_of_circle As Integer
'Global old_last_area_of_circle As Integer
'Global temp_last_area_of_circle As Integer
'Global last_area_of_circle_for_aid As Integer
'Global last_condition_area_of_circle As Integer
'area_of_polygon '45
'Global last_area_of_polygon As Integer
'Global old_last_area_of_polygon As Integer
'Global last_area_of_polygon_for_aid As Integer
'Global temp_last_area_of_polygon As Integer
'Global last_condition_area_of_polygon As Integer
'area_of_fan '46
Type distance_of_paral_line_data0_type
paral_no As Integer
lv_no As Integer
value As String
record As record_data_type
End Type
Type distance_of_paral_line_data_type
data(8) As distance_of_paral_line_data0_type
End Type
Global Ddistance_of_paral_line() As distance_of_paral_line_data_type
Type distance_of_point_line_data0_type
point_no  As Integer
line_no As Integer
dis_v As String
record As record_data_type
End Type
Type distance_of_point_line_data_type
data(8) As distance_of_point_line_data0_type
record_ As record_type
End Type
Global Ddistance_of_point_line() As distance_of_point_line_data_type
Type area_of_fan_data_type
poi(2) As Integer
record As record_data_type
'record_0 As record_type
value As String
value_ As String
End Type
Global area_of_fan_data_0 As area_of_fan_data_type
Type area_of_fan_type
data(8) As area_of_fan_data_type
 record_  As record_type
 End Type
Global Area_of_fan() As area_of_fan_type
Type con_area_of_fan_type
data(8) As area_of_fan_data_type
 'record_  As record_type
 End Type
Global con_Area_of_fan(3) As con_area_of_fan_type '½áÂÛ
'Global last_area_of_fan As Integer
'Global old_last_area_of_fan As Integer
'Global last_area_of_fan_for_aid As Integer
'Global temp_last_area_of_fan As Integer
'Global last_condition_area_of_fan As Integer
'sides_length_of_triangle47
Type sides_length_of_triangle_data_type
triangle As Integer
record As record_data_type
'record_0 As record_type
value As String
value_ As String
End Type
Global sides_length_of_triangle_data_0 As sides_length_of_triangle_data_type
Type sides_length_of_triangle_type
data(8)  As sides_length_of_triangle_data_type
 record_  As record_type
 End Type
Global Sides_length_of_triangle() As sides_length_of_triangle_type
Type con_sides_length_of_triangle_type
data(8)  As sides_length_of_triangle_data_type
 'record_  As record_type
 End Type
Global con_Sides_length_of_triangle(3) As con_sides_length_of_triangle_type '½áÂÛ
'Global last_sides_length_of_triangle As Integer
'Global old_last_sides_length_of_triangle As Integer
'Global last_sides_length_of_triangle_for_aid As Integer
'Global temp_last_sides_length_of_triangle As Integer
'Global last_condition_sides_length_of_triangle
'sides_length_of_circle48
Type sides_length_of_circle_data_type
circ As Integer
record As record_data_type
'record_0 As record_type
value As String
value_ As String
End Type
Global sides_length_of_circle_data_0 As sides_length_of_circle_data_type
Type sides_length_of_circle_type
data(8) As sides_length_of_circle_data_type
'Circ As Integer
'record As record_type
record_  As record_type
End Type
Global Sides_length_of_circle() As sides_length_of_circle_type
Type con_sides_length_of_circle_type
data(8) As sides_length_of_circle_data_type
'Circ As Integer
'record As record_type
'record_  As record_type
End Type
Global con_Sides_length_of_circle(3) As con_sides_length_of_circle_type '½áÂÛ
'Global last_sides_length_of_circle As Integer
'Global old_last_sides_length_of_circle As Integer
'Global last_sides_length_of_circle_for_aid As Integer
'Global temp_last_sides_length_of_circle As Integer
'Global last_condition_sides_length_of_circle As Integer
'verti_mid_line  '49
Type verti_mid_line_data0_type
 poi(2) As Integer
  n(2) As Integer
   line_no(1) As Integer
    mid_point_no  As Integer
     verti_no As Integer
End Type
Type verti_mid_line_data_type
 data0 As verti_mid_line_data0_type
      record As record_data_type
End Type
Type verti_mid_line_type
data(8) As verti_mid_line_data_type
 record_ As record_type
End Type
Global verti_mid_line_data0 As verti_mid_line_data0_type
Global verti_mid_line() As verti_mid_line_type
Type con_verti_mid_line_type
data(8) As verti_mid_line_data_type
'old_data As verti_mid_line_data_type
End Type
Global con_verti_mid_line(3) As con_verti_mid_line_type
'Global last_verti_mid_line(1) As Integer
'Global old_last_verti_mid_line As Integer
'Global temp_last_verti_mid_line As Integer
'Global last_verti_mid_line_for_aid(1) As Integer
'Global last_condition_verti_mid_line As Integer
'squ_sum 50
Type squ_sum_data_type
 poi(5) As Integer
 record As record_data_type
 'record_0 As record_type
 End Type
Type squ_sum_type
data(8) As squ_sum_data_type
 'poi(5) As Integer
 'record As record_type
 record_  As record_type
 End Type
 Global Squ_sum() As squ_sum_type
Type con_squ_sum_type
data(8) As squ_sum_data_type
 End Type
 Global con_con_squ_sum(3) As squ_sum_type
 'Global last_squ_sum As Integer
 'Global old_last_squ_sum As Integer
 'Global last_squ_sum_for_aid As Integer
 'Global temp_last_squ_sum As Integer
 'Global last_condition_squ_sum As Integer
 '****************
'¸¨Öúµã
'**********************************************
Type aid_point_data0_0
 triA(1) As temp_triangle_data_type
End Type
Type aid_point_data_type
 data(8) As aid_point_data0_0 ' Á½Èý½ÇÐÎ,ÒÔ±ãÏàÍ¬,Ò»½Ç»¥²¹
End Type
Global aid_point_data1() As aid_point_data_type
'Type aid_point_data2_type
' data(8) As aid_point_data0_0 ' Á½Èý½ÇÐÎ,ÒÔ±ãÏàÍ¬,Ò»½Ç»¥²¹
'End Type
Global aid_point_data2() As aid_point_data_type
Global aid_point_data3() As aid_point_data_type
'*******************

 '*********************
'Type circle_for_move_type
'center As Integer
'center_coord As POINTAPI
'circ As Integer
'radii As Long
'End Type
Type line_for_move_type
line_no As Integer
poi(1) As Integer
coord(1) As POINTAPI
r As Single
End Type

Global line_for_move As line_for_move_type
Global prove_by_hand_no(50) As Integer
Global prove_by_hand_type(50) As Integer
Global last_prove_by_hand_no As Integer
Global next_step_of_profe As Boolean
Global con_line1(8) As line_type
Global last_con_line1 As Integer
Global m_lin_0 As line_type '¹©³õÊ¼»¯
Global input_char_info As String
'Global theorem_text(1 To 256) As String
Global input_to_theorem_text(30) As String
Global run_type As Byte
Global run_type_1 As Byte
Global condition_no As Integer '¼ÇÂ¼ÊäÈëµÄ×Ö·ûÔÚwenti_condÖÐµÄÎ»ÖÃ
Global modify_condition_no As Integer
Global modify_icon_x As Integer
Global modify_icon_y As Integer
Global modify_icon_fontsize As Integer
Global global_icon_char As String * 1
Global problem_text As String * 255
'Global old_wenti_no As Integer
Global Const input_wenti = True
Global Const tuili = False '
Global input_p1%, input_p2% '¼ÇÂ¼ÒÑÖª£¬ÇóÖ¤£¬Ö¤Ã÷µÄÎ»ÖÃ
'*********************************
'Global circle_for_move As circle_for_move_type
Global measur_step  As Integer
'****************************************************************
Global display_type As Byte '  ¼ÇÂ¼ÏÔÊ¾µÄ·½Ê½
Global Const auto_display = 0
Global Const step_display = 1
Global code As Integer
Global using_area_th As Byte '0 ²»×÷Ãæ»ýÍÆÀí,1 Ãæ»ý,2Ãæ»ý±È
'*****************************************************************

'*****************************************
'  ÍÆÀí¼ÇÂ¼ÀàÐÍ
'*******************************************
Global prove_result As Byte
Global Const no_thing = 0
Global Const condition = 1
Global Const conclusion = 2
Global Const fill = 3
Global Const dot_mode = 4
Global Const aid_condition = 5
Global Const condition_and_project_line = 6 + 1
Global Const aid_condition_and_project_line = 5 + 1
Global Const display = True
Global Const delete = False
Global Const no_display = False
Global Const point3_on_line_ = 1
Global Const point4_on_circle_ = 9
Global Const Dangle_ = 88 'dpoint3 = 3
Global Const angle2_right = 4
Global Const relation_ = 7
Global Const v_relation_ = 131
Global Const dpoint_pair_ = 8
Global Const midpoint_ = 9
Global Const total_equal_triangle_ = 10
Global Const similar_triangle_ = 11
Global Const angle_ = 12
Global Const arc_ = 13
Global Const eline_ = 14
Global Const function_of_angle_ = 15
Global Const verti_mid_ = 16
Global Const inter_point_Dline_Dline = 17
Global Const Rangle_ = 18
Global Const item0_ = 19
Global Const tixing_ = 20
Global Const parallelogram_ = 110
Global Const Rtriangle_ = 21
Global Const Squsum = 23
Global Const triangle_ = 24
Global Const angle_value_ = 25
Global Const eangle_ = 26
Global Const two_angle_180 = 27
Global Const angle_relation_ = 28
Global Const two_angle_value_sum_ = 29
Global Const three_angle_value_ = 30
Global Const point3_on_circle = 32
Global Const distance_of_point_line_ = 33
Global Const distance_of_paral_line_ = 34
Global Const area_of_polygon_ = 35
Global Const area_of_triangle_ = 36
Global Const sp_polygon4_ = 37
Global Const icon_ = 38
Global Const inform_ = 39
Global Const TE_triangle_ssnnrn = 40
Global Const TE_triangle_nssnrn = 41
Global Const TE_triangle_nssnnr = 42
Global Const TE_triangle_snsrnn = 43
Global Const TE_triangle_snsnnr = 44
Global Const Dline_ = 45
Global Const V_line_value_ = 46
Global Const V_two_line_time_value_ = 47
Global Const S_triangle_sssnnn = 51
Global Const S_triangle_ssnnna = 52
Global Const S_triangle_snsnan = 53
Global Const S_triangle_nssann = 54
Global Const S_triangle_snnnaa = 55
Global Const S_triangle_nsnana = 56
Global Const S_triangle_nnsaan = 57
Global Const S_triangle_ssnrnn = 58
Global Const S_triangle_ssnnrn = 59
Global Const S_triangle_nssnrn = 60
Global Const S_triangle_nssnnr = 61
Global Const S_triangle_snsrnn = 62
Global Const S_triangle_snsnnr = 63
Global Const add_condition_ = 69
Global Const line3_value_ = 70
Global Const line_value_ = 71
Global Const angle3_value_ = 72
Global Const Two_angle_value_ = 73
Global Const arc_value_ = 74
Global Const equal_arc_ = 75
Global Const ratio_of_arc_ = 76
Global Const angle_less_angle_ = 77
Global Const line_less_line_ = 78
Global Const line_less_line2_ = 79
Global Const line2_less_line2_ = 80
Global Const tangent_line_ = 81
Global Const two_line_value_ = 82
Global Const mid_point_line_ = 83
Global Const area_relation_ = 84
Global Const reduce_angle3_value_ = 85
Global Const two_area_of_element_value_ = 86
Global Const two_point_conset_ = 87
Global Const general_string_ = 256
Global Const common_tangent_ = 88
Global Const epolygon_ = 89
Global Const general_angle_string_ = 90
Global Const point_type_ = 91
Global Const string_value_ = 92
Global Const line_from_two_point_ = 93
Global Const pseudo_total_equal_triangle_ = 94
Global Const pseudo_similar_triangle_ = 95
Global Const two_order_equation_ = 96
'Èý½ÇÐÎµÄÀàÐÍ
Global Const equal_sides = 101
Global Const equal_side_triangle_ = 102
Global Const equal_side_right_triangle_ = 103
Global Const equal_side_1 = 104
Global Const equal_side_2 = 105
Global Const right_angle = 106
Global Const right_equal_side_0 = 107
Global Const right_equal_side_1 = 111
Global Const right_equal_side_2 = 109
Global Const equal_side_tixing_ = 108
Global Const circle_ = 253
Global Const long_squre_ = 113
Global Const rhombus_ = 112
Global Const Squre = 200
Global Const area_of_element_ = 114
Global Const area_of_circle_ = 115
Global Const length_of_polygon_ = 116
Global Const area_of_fan_ = 117
Global Const sides_length_of_triangle_ = 118
Global Const sides_length_of_circle_ = 119
Global Const verti_mid_line_ = 120
Global Const tangent_circle_ = 121
Global Const point_in_mid_verti_line_ = 122
Global Const point_ = 251
Global Const same_three_lines_ = 124
Global Const polygon_ = 125
Global Const line_ = 252
Global Const wenti_cond_ = 254
Global Const aid_line_ = 203
Global Const relation_string_ = 128
Global Const tri_function_ = 129
Global Const total_angle_ = 130
Global Const equation_ = 130
Global Const plane_ = 131
Global Const related_line_ = 500
Global Const wait_for_prove = 10000
Global Const sin_sig = -1 '"$"
Global Const cos_sig = -2 '"&"
Global Const tan_sig = -3 '"`"
Global Const tg_sig = -3 '"`"
Global Const ctg_sig = -4 '"\"
Global Const ctan_sig = -4 '"\"
Global Const item_sig = -5
Global Const angle_sig = -6
'*********************************
Global Const IO_yes = True
Global Const IO_no = False
'********************************************
'  ËùÓÃ¶¨ÀíÀàÐÍ
'*********************************************
Global Const T1001 = 1001
Global Const T1002 = 1002
Global Const T1003 = 1003
Global Const T1004 = 1004
Global Const T1005 = 1005
'*********************************************
'************************************************
Global conclusion_no_wenti As Integer '½áÂÛÓï¾äºÅ
Type display_string_type
 display_record_type As Byte
  display_record_no As Integer
'   aid_string_no As Integer
    conclusion_or_condition As Byte
     condition_data As condition_data_type
      reduce_level As Byte
       is_same_theorem As Byte
End Type
Global inform_type As Byte
Global display_string() As display_string_type
Global display_no As Integer
Type one_triangle_data_type
triangle As Integer
record As record_data_type
direction As Integer
End Type
Global one_triangle_data_0 As one_triangle_data_type
Type one_triangle_type
data(8) As one_triangle_data_type
record_ As record_type
End Type
Type con_one_triangle_type
data(8) As one_triangle_data_type
record_ As record_type
End Type
'********************************************
'****************************************
'*******************************************
Global con_equal_side_triangle(3) As con_one_triangle_type
Global con_equal_side_right_triangle(3) As con_one_triangle_type
Global equal_side_triangle() As one_triangle_type
'Global last_equal_side_triangle As Integer
Global equal_side_right_triangle() As one_triangle_type
'Global last_equal_side_right_triangle As Integer
'Global old_last_equal_side_right_triangle As Integer
'Global last_equal_side_right_triangle_for_aid As Integer
'Global old_last_equal_side_right_triangle_for_aid As Integer
Type conclusion_data_type
no(1) As Integer
branch As Integer
ty As Integer
wenti_no As Integer
End Type
Global conclusion_data(3) As conclusion_data_type ' As Integer '( x,1) 'º¬Î´ÖªÊýµÄ½áÂÛ
Global last_condition_record_no As Integer
'¼ÇÂ¼×îºóÒ»¸öÌõ¼þ

Global display_theorem As Boolean

'****************************************************
Type record_for_trans_type
 last_trans_to As Integer
  record() As record_type0
End Type
Type item0_data_type
para(1) As String
poi(5) As Integer 'µã
n(5) As Integer
line_no(2) As Integer 'Ïß¶Î
sig As String * 1 'ÔËËã·ûºÅ
diff_type As Byte '²îÐÍ¿É»¯³ÉÁ½Ïî
diff_para(1) As String 'ÏµÊý
diff_it(1) As Integer '
is_const As Byte
big_or_smamll As Boolean
record_for_initial As condition_data_type '¼ÇÂ¼¹²ÏßµÄÌõ¼þ
record_for_value As record_data_type
record_for_diff As record_type0
record_for_trans As record_for_trans_type
value As String
index(3) As Integer
no_reduce As Boolean
conclusion_no As Integer
'±íÊ¾´ËÏîÀ´×Ô½áÂÛ
End Type
Type item0_type
record As record_type
data(8) As item0_data_type
End Type
Global item0() As item0_type
Type element_data_type
poi(1) As Integer
End Type
Global last_condition_item0 As Integer
Type function_data_type
variant_data As element_data_type
function_data As element_data_type
End Type
Global is_set_function_data As Byte '=1 ÉèÖÃ×Ô±äÁ¿=2,ÉèÖÃº¯ÊýÁ¿=3ÉèÖÃÍê³É
'º¯Êý¹ØÏµ
Global function_data As function_data_type
'**********************************************************
Type four_sides_fig_data_type
poi(3) As Integer
index(3) As Integer
triA(1) As Integer
End Type
Type four_sides_fig_type
data(8) As four_sides_fig_data_type
End Type
Global four_sides_fig() As four_sides_fig_type
Type Dline1_data_type
poi(1) As Integer
record As record_data_type
'record_0 As record_type
End Type
Global Dline1_data_0 As Dline1_data_type
Type Dline1_type
data(8) As Dline1_data_type
record_ As record_type
End Type
Global Dline1() As Dline1_type
'Global last_dline1 As Integer
'Global old_last_dline1 As Integer
'Global last_aid_dline1 As Integer
'****************************************************88
Type Dangle_data_type
angle As Integer
record As record_data_type
'record_0 As record_type
End Type
Global Dangle_data_0 As Dangle_data_type
Type Dangle_type
data(8) As Dangle_data_type
'angle As Integer
'record As record_type
record_ As record_type
End Type
Type function_of_angle_data_type
angle_no As Integer
value(3) As String
record As record_data_type
End Type
Global function_of_angle_data_0 As function_of_angle_data_type
Type function_of_angle_type
data(8) As function_of_angle_data_type
record_  As record_type
End Type
Type con_function_of_angle_type
data(8) As function_of_angle_data_type
End Type
Global function_of_angle() As function_of_angle_type
Global con_function_of_angle(3) As con_function_of_angle_type
Global v_coordinate_system_no(1) As Integer
Global Dangle() As Dangle_type
'***************************************************
Type temp_mid_point_type
poi(2) As Integer
End Type
'*********************************************
Type special_angle_type
angle As Integer
vulue As String
record As record_type
record_0 As record_type
End Type
'***************************************
Global Dangle30() As special_angle_type
Global last_Dangle30 As Integer
Global old_last_Dangle30 As Integer
Global last_aid_Dangle30 As Integer
'**************************************
Global Dangle45() As special_angle_type
Global last_Dangle45 As Integer
Global old_last_Dangle45 As Integer
Global last_aid_Dangle45 As Integer
'**************************************
Global Dangle60() As special_angle_type
Global last_Dangle60 As Integer
Global old_last_Dangle60 As Integer
Global last_aid_Dangle60 As Integer
'**************************************
Global Dangle120() As special_angle_type
Global last_Dangle120 As Integer
Global old_last_Dangle120 As Integer
Global last_aid_Dangle120 As Integer
'**************************************
Global Dangle135() As special_angle_type
Global last_Dangle135 As Integer
Global old_last_Dangle135 As Integer
Global last_aid_Dangle135 As Integer
'**************************************
Global Dangle150() As special_angle_type
Global last_Dangle150 As Integer
Global old_last_Dangle150 As Integer
Global last_aid_Dangle150 As Integer
'**************************************
'****************************************************
Type right_triangle_type
poi(2) As Integer
record As record_type
record_0 As record_type
End Type
Global right_triangle() As right_triangle_type
'*******************************************
Type right_angle_type
record As record_type
record_0 As record_type
 angle As Integer
End Type

'**********************************************************
Type two_angle_type
record As record_type
record_0 As record_type
angle(2) As Integer 'angle(2) ±íÊ¾ºÍ
sum_direction As Boolean
poi(1) As Integer
End Type

'*********************************
Global comp_num1(16), comp_num2(16) As Integer
'***************************************************************
Type hotpoint_of_theorem_type
'display_string_no As Byte
'ÏàÓ¦µÄÖ¤Ã÷ÏÔÊ¾ºÅ
theorem_no As Integer
cond_no As Integer
'hot_string As String
'ÈÈµãÓï¾ä
X(1) As Integer
Y(1) As Integer
'ÉèÖÃÈÈµãµÄ·¶Î§
End Type


'Global Hotpoint_of_theorem() As hotpoint_of_theorem_type
'Global Last_hotpoint_of_theorem As Integer
Global Hotpoint_of_theorem1() As hotpoint_of_theorem_type
Global Last_hotpoint_of_theorem1 As Integer


Public Sub set_depend_condition(ty As Integer, n%)
Dim no%, i%, p%
Select Case ty
Case area_of_element_
  area_of_element(n%).record_.is_depend = 1
   no% = Abs(area_of_element(n%).record_.display_no)
Case sides_length_of_triangle_
  Sides_length_of_triangle(n%).record_.is_depend = 1
   no% = Abs(Sides_length_of_triangle(n%).record_.display_no)
Case area_of_circle_
  area_of_circle(n%).record_.is_depend = 1
   no% = Abs(area_of_circle(n%).record_.display_no)
Case sides_length_of_circle_
  Sides_length_of_circle(n%).record_.is_depend = 1
   no% = Abs(Sides_length_of_circle(n%).record_.display_no)
Case area_of_fan_
   Area_of_fan(n%).record_.is_depend = 1
   no% = Abs(Area_of_fan(n%).record_.display_no)
Case two_line_value_
   two_line_value(n%).record_.is_depend = 1
   no% = Abs(two_line_value(n%).record_.display_no)
Case rhombus_
  rhombus(n%).record_.is_depend = 1
   no% = Abs(rhombus(n%).record_.display_no)
Case long_squre_
  Dlong_squre(n%).record_.is_depend = 1
   no% = Abs(Dlong_squre(n%).record_.display_no)
Case Squre
  Dsqure(n%).record_.is_depend = 1
   no% = Abs(Dsqure(n%).record_.display_no)
Case equation_
  equation(n%).record_.is_depend = 1
   no% = Abs(equation(n%).record_.display_no)
Case general_string_
  general_string(n%).record_.is_depend = 1
   no% = Abs(general_string(n%).record_.display_no)
Case tangent_line_
  tangent_line(n%).record_.is_depend = 1
   no% = Abs(tangent_line(n%).record_.display_no)
Case equal_side_triangle_
  equal_side_right_triangle(n%).record_.is_depend = 1
   no% = Abs(equal_side_triangle(n%).record_.display_no)
Case epolygon_
  epolygon(n%).record_.is_depend = 1
   no% = Abs(epolygon(n%).record_.display_no)
Case equal_arc_
  equal_arc(n%).record_.is_depend = 1
   no% = Abs(equal_arc(n%).record_.display_no)
Case area_relation_
 Darea_relation(n%).record_.is_depend = 1
   no% = Abs(Darea_relation(n%).record_.display_no)
'**********************************************************
Case angle3_value_
  angle3_value(n%).record_.is_depend = 1
   no% = Abs(angle3_value(n%).record_.display_no)
'**********************************************************
Case Dangle_
  Dangle(n%).record_.is_depend = 1
   no% = Abs(Dangle(n%).record_.display_no)
'************************************************************
Case Dline_
  Dline1(n%).record_.is_depend = 1
   no% = Abs(Dline1(n%).record_.display_no)
'***************************************************
Case dpoint_pair_
  Ddpoint_pair(n%).record_.is_depend = 1
   no% = Abs(Ddpoint_pair(n%).record_.display_no)
'*******************************************************
Case eline_
  Deline(n%).record_.is_depend = 1
   no% = Abs(Deline(n%).record_.display_no)
Case midpoint_ 'equal_segment_on_line
  Dmid_point(n%).record_.is_depend = 1
   no% = Abs(Dmid_point(n%).record_.display_no)
'***********************************************************
Case mid_point_line_
  mid_point_line(n%).record_.is_depend = 1
   no% = Abs(mid_point_line(n%).record_.display_no)
Case line_value_
  line_value(n%).record_.is_depend = 1
   no% = Abs(line_value(n%).record_.display_no)
'**********************************************************
Case line3_value_
  line3_value(n%).record_.is_depend = 1
   no% = Abs(line3_value(n%).record_.display_no)
'*****************************************************
Case paral_
  Dparal(n%).record_.is_depend = 1
   no% = Abs(Dparal(n%).record_.display_no)
'*********************************************
Case parallelogram_
  Dparallelogram(n%).record_.is_depend = 1
   no% = Abs(Dparallelogram(n%).record_.display_no)
'***************************************************
Case point3_on_line_
  three_point_on_line(n%).record_.is_depend = 1
   no% = Abs(three_point_on_line(n%).record_.display_no)
'****************************************************************
Case point4_on_circle_
  four_point_on_circle(n%).record_.is_depend = 1
   no% = Abs(four_point_on_circle(n%).record_.display_no)
'*************************************************************
Case relation_
  Drelation(n%).record_.is_depend = 1
   no% = Abs(Drelation(n%).record_.display_no)
'*************************************************************
Case v_relation_
  Drelation(n%).record_.is_depend = 1
   no% = Abs(Drelation(n%).record_.display_no)
'*************************************************************
Case v_relation_
  Drelation(n%).record_.is_depend = 1
   no% = Abs(Drelation(n%).record_.display_no)
'****************************************************************
Case similar_triangle_
  Dsimilar_triangle(n%).record_.is_depend = 1
   no% = Abs(Dsimilar_triangle(n%).record_.display_no)
'****************************************************
Case tri_function_
  tri_function(n%).record_.is_depend = 1
   no% = Abs(tri_function(n%).record_.display_no)
'*************************************************************
Case total_equal_triangle_
  Dtotal_equal_triangle(n%).record_.is_depend = 1
   no% = Abs(Dtotal_equal_triangle(n%).record_.display_no)
Case verti_
  Dverti(n%).record_.is_depend = 1
'*********************************************************
   no% = Abs(Dverti(n%).record_.display_no)
Case verti_mid_line_
  verti_mid_line(n%).record_.is_depend = 1
   no% = Abs(verti_mid_line(n%).record_.display_no)
'*********************************************************
Case 0
  no% = 0
End Select
    If no% > 0 And no% <= C_display_wenti.m_last_conclusion Then '
    Call C_display_wenti.set_m_depend_no(no%)
    For i% = 0 To 10
     p% = C_display_wenti.m_point_no(no%, i%)
     If p% > 0 Then
        If C_display_wenti.m_condition(no%, i%) >= "A" And _
              C_display_wenti.m_condition(no%, i%) <= "Z" Then
         Call set_depend_from_point(p%)
        End If
     End If
    Next i%
    End If
End Sub
Public Sub record_no(ByVal ty As Integer, ByVal n%, re As total_record_type, _
   init_display As Boolean, operat_type As Byte, remove_no%)
Dim i%
'ty Êý¾ÝÀàÐÍn¡¡Êý¾Ýre·µ»ØµÄ¼ÇÂ¼
'Global Const no_thing = 0
'Global Const point1_on_line = 1
'Global Const point2_on_line = 2
'Global Const point3_on_line = 3
'Global Const point1_on_circle = 4
'Global Const point2_on_circle =5
'Global Const point3_on_circle = 6
'Global Const point4_on_circle = 7
'Global Const dpoint3 = 8
'Global Const dpoint4 = 9
'Global Const paral = 10
'Global Const verti = 11
'Global Const relation = 12
'Global Const eangle = 13
'Global Const eline = 14
'Global Const mid_point_of_line = 15
'Global Const verti_mid = 16
'Global Const inter_point_Dline_Dline = 17
'Global Const inter_point_Dline_Dcircle = 18
'Global Const inter_point_Dcircle_Dcircle = 19
'Call simple_record(ty, n%, re)
' Dim temp_record  As record_type
Select Case ty
'******************************************
Case new_point_
If operat_type = 0 Then
  re.record_data = new_point(n%).data(0).record
  re.record_ = new_point(n%).record_
  If init_display = True And new_point(n%).record_.display_no > 0 Then
   new_point(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 'new_point(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 new_point(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   new_point(n%).data(0).record.data0.condition_data.condition_no = _
      new_point(n%).data(0).record.data0.condition_data.condition_no - 1
    For i% = remove_no% To new_point(n%).data(0).record.data0.condition_data.condition_no
     new_point(n%).data(0).record.data0.condition_data.condition(i%) = _
     new_point(n%).data(0).record.data0.condition_data.condition(i% + 1)
    Next i%
End If
'************************************************************
Case item0_
If operat_type = 0 Then
  re.record_data = item0(n%).data(0).record_for_value
  re.record_ = item0(n%).record
  If init_display = True And item0(n%).record.display_no > 0 Then
   item0(n%).record.display_no = 0
  End If
ElseIf operat_type = 1 Then
 'new_point(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 item0(n%).record = re.record_
End If
If remove_no% > 0 Then
   item0(n%).data(0).record_for_value.data0.condition_data.condition_no = _
      item0(n%).data(0).record_for_value.data0.condition_data.condition_no - 1
    For i% = remove_no% To item0(n%).data(0).record_for_value.data0.condition_data.condition_no
     item0(n%).data(0).record_for_value.data0.condition_data.condition(i%) = _
     item0(n%).data(0).record_for_value.data0.condition_data.condition(i% + 1)
    Next i%
End If
Case pseudo_similar_triangle_
 If operat_type = 0 Then
  re.record_data = pseudo_similar_triangle(n%).data(0).record
  re.record_ = pseudo_similar_triangle(n%).record_
   If init_display = True And pseudo_similar_triangle(n%).record_.display_no > 0 Then
    pseudo_similar_triangle(n%).record_.display_no = 0
   End If
 ElseIf operat_type = 1 Then
   pseudo_similar_triangle(n%).data(0).record.data0 = re.record_data.data0
 ElseIf operat_type = 2 Then
  pseudo_similar_triangle(n%).record_ = re.record_
 End If
If remove_no% > 0 Then
   pseudo_similar_triangle(n%).data(0).record.data0.condition_data.condition_no = _
      pseudo_similar_triangle(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To pseudo_similar_triangle(n%).data(0).record.data0.condition_data.condition_no
   pseudo_similar_triangle(n%).data(0).record.data0.condition_data.condition(i%) = _
     pseudo_similar_triangle(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'************************************************************************
Case pseudo_total_equal_triangle_
 If operat_type = 0 Then
  re.record_data = pseudo_total_equal_triangle(n%).data(0).record
  re.record_ = pseudo_total_equal_triangle(n%).record_
   If init_display = True And pseudo_total_equal_triangle(n%).record_.display_no > 0 Then
    pseudo_total_equal_triangle(n%).record_.display_no = 0
   End If
 ElseIf operat_type = 1 Then
  pseudo_total_equal_triangle(n%).data(0).record.data0 = re.record_data.data0
 ElseIf operat_type = 2 Then
  pseudo_total_equal_triangle(n%).record_ = re.record_
 End If
 If remove_no% > 0 Then
   pseudo_total_equal_triangle(n%).data(0).record.data0.condition_data.condition_no = _
      pseudo_total_equal_triangle(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To pseudo_total_equal_triangle(n%).data(0).record.data0.condition_data.condition_no
   pseudo_total_equal_triangle(n%).data(0).record.data0.condition_data.condition(i%) = _
     pseudo_total_equal_triangle(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
 End If
'*************************************************
Case V_line_value_
 If operat_type = 0 Then
  re.record_data = V_line_value(n%).data(0).record
  re.record_ = V_line_value(n%).record_
   If init_display = True And V_line_value(n%).record_.display_no > 0 Then
    V_line_value(n%).record_.display_no = 0
   End If
 ElseIf operat_type = 1 Then
  V_line_value(n%).data(0).record.data0 = re.record_data.data0
 ElseIf operat_type = 2 Then
  V_line_value(n%).record_ = re.record_
 End If
 If remove_no% > 0 Then
   V_line_value(n%).data(0).record.data0.condition_data.condition_no = _
      V_line_value(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To V_line_value(n%).data(0).record.data0.condition_data.condition_no
   V_line_value(n%).data(0).record.data0.condition_data.condition(i%) = _
     V_line_value(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
 End If
'*************************************************
Case sides_length_of_triangle_
If operat_type = 0 Then
 re.record_data = Sides_length_of_triangle(n%).data(0).record
 re.record_ = Sides_length_of_triangle(n%).record_
  If init_display = True And Sides_length_of_triangle(n%).record_.display_no > 0 Then
   Sides_length_of_triangle(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 Sides_length_of_triangle(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Sides_length_of_triangle(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Sides_length_of_triangle(n%).data(0).record.data0.condition_data.condition_no = _
      Sides_length_of_triangle(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Sides_length_of_triangle(n%).data(0).record.data0.condition_data.condition_no
   Sides_length_of_triangle(n%).data(0).record.data0.condition_data.condition(i%) = _
     Sides_length_of_triangle(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'*******************************************
'Case distance_of_paral_line
'If operat_type = 0 Then
' re.record_data = Ddistance_of_paral_line(n%).data(0).record
' re.record_ = Ddistance_of_paral_line(n%).record_
'  If init_display = True And Ddistance_of_paral_line(n%).record_.display_no > 0 Then
'   Ddistance_of_paral_line(n%).record_.display_no = 0
'  End If
'ElseIf operat_type = 1 Then
' Ddistance_of_paral_line(n%).data(0).record.data0 = re.record_data.data0
'ElseIf operat_type = 2 Then
' Ddistance_of_paral_line(n%).record_ = re.record_
'End If
'If remove_no% > 0 Then
'   Ddistance_of_paral_line(n%).data(0).record.data0.condition_data.condition_no = _
      Ddistance_of_paral_line(n%).data(0).record.data0.condition_data.condition_no - 1
'  For i% = remove_no% To Ddistance_of_paral_line(n%).data(0).record.data0.condition_data.condition_no
'   Ddistance_of_paral_line(n%).data(0).record.data0.condition_data.condition(i%) = _
'     Ddistance_of_paral_line(n%).data(0).record.data0.condition_data.condition(i% + 1)
'  Next i%
'End If
'****************************************************
'Case distance_of_paral_line
'If operat_type = 0 Then
' re.record_data = Ddistance_of_point_line(n%).data(0).record
' re.record_ = Ddistance_of_point_line(n%).record_
'  If init_display = True And Ddistance_of_point_line(n%).record_.display_no > 0 Then
'   Ddistance_of_point_line(n%).record_.display_no = 0
'  End If
'ElseIf operat_type = 1 Then
'   Ddistance_of_point_line(n%).data(0).record.data0 = re.record_data.data0
'ElseIf operat_type = 2 Then
'   Ddistance_of_point_line(n%).record_ = re.record_
'End If
'If remove_no% > 0 Then
'   Ddistance_of_point_line(n%).data(0).record.data0.condition_data.condition_no = _
'      Ddistance_of_point_line(n%).data(0).record.data0.condition_data.condition_no - 1
'  For i% = remove_no% To Ddistance_of_point_line(n%).data(0).record.data0.condition_data.condition_no
'   Ddistance_of_point_line(n%).data(0).record.data0.condition_data.condition(i%) = _
'     Ddistance_of_point_line(n%).data(0).record.data0.condition_data.condition(i% + 1)
'  Next i%
'End If
'**************************************************************
Case sides_length_of_circle_
If operat_type = 0 Then
 re.record_data = Sides_length_of_circle(n%).data(0).record
 re.record_ = Sides_length_of_circle(n%).record_
  If init_display = True And Sides_length_of_circle(n%).record_.display_no > 0 _
   Then
   Sides_length_of_circle(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 Sides_length_of_circle(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Sides_length_of_circle(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Sides_length_of_circle(n%).data(0).record.data0.condition_data.condition_no = _
      Sides_length_of_circle(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Sides_length_of_circle(n%).data(0).record.data0.condition_data.condition_no
   Sides_length_of_circle(n%).data(0).record.data0.condition_data.condition(i%) = _
     Sides_length_of_circle(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'*************************
Case area_of_circle_
If operat_type = 0 Then
 re.record_data = area_of_circle(n%).data(0).record
 re.record_ = area_of_circle(n%).record_
  If init_display = True And area_of_circle(n%).record_.display_no > 0 Then
   area_of_circle(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 area_of_circle(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 area_of_circle(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   area_of_circle(n%).data(0).record.data0.condition_data.condition_no = _
      area_of_circle(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To area_of_circle(n%).data(0).record.data0.condition_data.condition_no
   area_of_circle(n%).data(0).record.data0.condition_data.condition(i%) = _
     area_of_circle(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
 End If
'***************************
Case area_of_element_
If operat_type = 0 Then
 If n% < 0 Then
 re.record_data.data0.condition_data.condition_no = 2
 re.record_data.data0.condition_data.condition(1).ty = area_of_element_
 re.record_data.data0.condition_data.condition(1).no = -n%
 re.record_data.data0.condition_data.condition(2) = area_of_element(-n%).data(0).record.data0.condition_for_value_
 re.record_data.data0.theorem_no = 1
 Else
 re.record_data = area_of_element(n%).data(0).record
 re.record_ = area_of_element(n%).record_
  If init_display = True And area_of_element(n%).record_.display_no > 0 Then
   area_of_element(n%).record_.display_no = 0
  End If
 End If
ElseIf operat_type = 1 Then
 area_of_element(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 area_of_element(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   area_of_element(n%).data(0).record.data0.condition_data.condition_no = _
      area_of_element(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To area_of_element(n%).data(0).record.data0.condition_data.condition_no
   area_of_element(n%).data(0).record.data0.condition_data.condition(i%) = _
     area_of_element(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'********************************
Case two_area_of_element_value_
If operat_type = 0 Then
 re.record_data = two_area_of_element_value(n%).data(0).record
 re.record_ = two_area_of_element_value(n%).record_
  If init_display = True And two_area_of_element_value(n%).record_.display_no > 0 Then
   two_area_of_element_value(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 two_area_of_element_value(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 two_area_of_element_value(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   two_area_of_element_value(n%).data(0).record.data0.condition_data.condition_no = _
      two_area_of_element_value(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To two_area_of_element_value(n%).data(0).record.data0.condition_data.condition_no
   two_area_of_element_value(n%).data(0).record.data0.condition_data.condition(i%) = _
     two_area_of_element_value(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'******************************
Case area_of_fan_
If operat_type = 0 Then
 re.record_data = Area_of_fan(n%).data(0).record
 re.record_ = Area_of_fan(n%).record_
  If init_display = True And Area_of_fan(n%).record_.display_no > 0 Then
   Area_of_fan(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 Area_of_fan(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Area_of_fan(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Area_of_fan(n%).data(0).record.data0.condition_data.condition_no = _
      Area_of_fan(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Area_of_fan(n%).data(0).record.data0.condition_data.condition_no
   Area_of_fan(n%).data(0).record.data0.condition_data.condition(i%) = _
     Area_of_fan(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'*********************************
Case tixing_
If operat_type = 0 Then
 re.record_data = Dtixing(n%).data(0).record
 re.record_ = Dtixing(n%).record_
  If init_display = True And Dtixing(n%).record_.display_no > 0 Then
   Dtixing(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 Dtixing(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Dtixing(n%).record_ = re.record_
End If
If remove_no% > 0 Then
    Dtixing(n%).data(0).record.data0.condition_data.condition_no = _
       Dtixing(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Dtixing(n%).data(0).record.data0.condition_data.condition_no
    Dtixing(n%).data(0).record.data0.condition_data.condition(i%) = _
      Dtixing(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'********************************
Case rhombus_
If operat_type = 0 Then
 re.record_data = rhombus(n%).data(0).record
 re.record_ = rhombus(n%).record_
  If init_display = True And rhombus(n%).record_.display_no > 0 Then
   rhombus(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 rhombus(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 rhombus(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   rhombus(n%).data(0).record.data0.condition_data.condition_no = _
      rhombus(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To rhombus(n%).data(0).record.data0.condition_data.condition_no
   rhombus(n%).data(0).record.data0.condition_data.condition(i%) = _
     rhombus(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'*******************************
Case long_squre_
If operat_type = 0 Then
 re.record_data = Dlong_squre(n%).data(0).record
 re.record_ = Dlong_squre(n%).record_
  If init_display = True And Dlong_squre(n%).record_.display_no > 0 Then
   Dlong_squre(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 Dlong_squre(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Dlong_squre(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Dlong_squre(n%).data(0).record.data0.condition_data.condition_no = _
      Dlong_squre(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Dlong_squre(n%).data(0).record.data0.condition_data.condition_no
   Dlong_squre(n%).data(0).record.data0.condition_data.condition(i%) = _
     Dlong_squre(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'***********************************
Case Squre
If operat_type = 0 Then
 re.record_data = Dsqure(n%).data(0).record
 re.record_ = Dsqure(n%).record_
  If init_display = True And Dsqure(n%).record_.display_no > 0 Then
   Dsqure(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 Dsqure(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Dsqure(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Dsqure(n%).data(0).record.data0.condition_data.condition_no = _
      Dsqure(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Dsqure(n%).data(0).record.data0.condition_data.condition_no
   Dsqure(n%).data(0).record.data0.condition_data.condition(i%) = _
     Dsqure(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'**************************************
Case equation_
If operat_type = 0 Then
 re.record_data = equation(n%).data(0).record
 re.record_ = equation(n%).record_
  If init_display = True And equation(n%).record_.display_no > 0 Then
   equation(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 equation(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 equation(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   equation(n%).data(0).record.data0.condition_data.condition_no = _
      equation(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To equation(n%).data(0).record.data0.condition_data.condition_no
   equation(n%).data(0).record.data0.condition_data.condition(i%) = _
     equation(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'Case equal_side_tixing_
' If operat_type = 0 Then
' re.record_data = Dequal_side_tixing(n%).data(0).record
' re.record_ = Dequal_side_tixing(n%).record_
'   If init_display = True And Dequal_side_tixing(n%).record_.display_no > 0 Then
'    Dequal_side_tixing(n%).record_.display_no = 0
'   End If
' ElseIf operat_type = 1 Then
'  Dequal_side_tixing(n%).data(0).record.data0 = re.record_data.data0
' ElseIf operat_type = 2 Then
'  Dequal_side_tixing(n%).record_ = re.record_
' End If
'************************************
Case equal_side_triangle_
 If operat_type = 0 Then
 re.record_data = equal_side_triangle(n%).data(0).record
 re.record_ = equal_side_triangle(n%).record_
   If init_display = True And equal_side_triangle(n%).record_.display_no > 0 Then
    equal_side_triangle(n%).record_.display_no = 0
   End If
 ElseIf operat_type = 1 Then
 equal_side_triangle(n%).data(0).record.data0 = re.record_data.data0
 ElseIf operat_type = 2 Then
 equal_side_triangle(n%).record_ = re.record_
 End If
If remove_no% > 0 Then
   equal_side_triangle(n%).data(0).record.data0.condition_data.condition_no = _
      equal_side_triangle(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To equal_side_triangle(n%).data(0).record.data0.condition_data.condition_no
   equal_side_triangle(n%).data(0).record.data0.condition_data.condition(i%) = _
     equal_side_triangle(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'***********************************
Case equal_side_right_triangle_
 If operat_type = 0 Then
 re.record_data = equal_side_right_triangle(n%).data(0).record
 re.record_ = equal_side_right_triangle(n%).record_
   If init_display = True And equal_side_right_triangle(n%).record_.display_no > 0 Then
    equal_side_right_triangle(n%).record_.display_no = 0
   End If
 ElseIf operat_type = 1 Then
  equal_side_right_triangle(n%).data(0).record.data0 = re.record_data.data0
 ElseIf operat_type = 2 Then
  equal_side_right_triangle(n%).record_ = re.record_
 End If
If remove_no% > 0 Then
   equal_side_right_triangle(n%).data(0).record.data0.condition_data.condition_no = _
      equal_side_right_triangle(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To equal_side_right_triangle(n%).data(0).record.data0.condition_data.condition_no
   equal_side_right_triangle(n%).data(0).record.data0.condition_data.condition(i%) = _
     equal_side_right_triangle(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'***********************************
Case tangent_line_
 If operat_type = 0 Then
 re.record_data = tangent_line(n%).data(0).record
 re.record_ = tangent_line(n%).record_
   If init_display = True And tangent_line(n%).record_.display_no > 0 Then
    tangent_line(n%).record_.display_no = 0
   End If
 ElseIf operat_type = 1 Then
  tangent_line(n%).data(0).record.data0 = re.record_data.data0
 ElseIf operat_type = 2 Then
  tangent_line(n%).record_ = re.record_
 End If
If remove_no% > 0 Then
   tangent_line(n%).data(0).record.data0.condition_data.condition_no = _
      tangent_line(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To tangent_line(n%).data(0).record.data0.condition_data.condition_no
   tangent_line(n%).data(0).record.data0.condition_data.condition(i%) = _
     tangent_line(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
 End If
'**************************************
Case epolygon_
 If operat_type = 0 Then
 re.record_data = epolygon(n%).data(0).record
 re.record_ = epolygon(n%).record_
   If init_display And epolygon(n%).record_.display_no > 0 Then
    epolygon(n%).record_.display_no = 0
   End If
 ElseIf operat_type = 1 Then
  epolygon(n%).data(0).record.data0 = re.record_data.data0
 ElseIf operat_type = 2 Then
  epolygon(n%).record_ = re.record_
 End If
If remove_no% > 0 Then
   epolygon(n%).data(0).record.data0.condition_data.condition_no = _
      epolygon(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To epolygon(n%).data(0).record.data0.condition_data.condition_no
   epolygon(n%).data(0).record.data0.condition_data.condition(i%) = _
     epolygon(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'**********************************
Case equal_arc_
 If operat_type = 0 Then
 re.record_data = equal_arc(n%).data(0).record
 re.record_ = equal_arc(n%).record_
   If init_display And equal_arc(n%).record_.display_no > 0 Then
    equal_arc(n%).record_.display_no = 0
   End If
 ElseIf operat_type = 1 Then
  equal_arc(n%).data(0).record.data0 = re.record_data.data0
 ElseIf operat_type = 2 Then
  equal_arc(n%).record_ = re.record_
 End If
If remove_no% > 0 Then
   equal_arc(n%).data(0).record.data0.condition_data.condition_no = _
      equal_arc(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To equal_arc(n%).data(0).record.data0.condition_data.condition_no
   equal_arc(n%).data(0).record.data0.condition_data.condition(i%) = _
     equal_arc(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'************************************
Case general_string_
 If operat_type = 0 Then
 re.record_data = general_string(n%).data(0).record
 re.record_ = general_string(n%).record_
   If init_display And general_string(n%).record_.display_no > 0 Then
    general_string(n%).record_.display_no = 0
   End If
 ElseIf operat_type = 1 Then
  general_string(n%).data(0).record.data0 = re.record_data.data0
 ElseIf operat_type = 2 Then
  general_string(n%).record_ = re.record_
 End If
If remove_no% > 0 Then
   general_string(n%).data(0).record.data0.condition_data.condition_no = _
      general_string(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To general_string(n%).data(0).record.data0.condition_data.condition_no
   general_string(n%).data(0).record.data0.condition_data.condition(i%) = _
     general_string(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'Case equal_area_triangle_
'If operat_type = 0 Then
' re.record_data = equal_area_triangle(n%).data(0).record
' re.record_ = equal_area_triangle(n%).record_
'  If init_display And equal_area_triangle(n%).record_.display_no > 0 Then
'   equal_area_triangle(n%).record_.display_no = 0
'  End If
'ElseIf operat_type = 1 Then
' equal_area_triangle(n%).data(0).record.data0 = re.record_data.data0
'ElseIf operat_type = 2 Then
' equal_area_triangle(n%).record_ = re.record_
'End If
'*************************************
Case area_relation_
If operat_type = 0 Then
 re.record_data = Darea_relation(n%).data(0).record
 re.record_ = Darea_relation(n%).record_
 If init_display And Darea_relation(n%).record_.display_no > 0 Then
  Darea_relation(n%).record_.display_no = 0
 End If
ElseIf operat_type = 1 Then
 Darea_relation(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Darea_relation(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Darea_relation(n%).data(0).record.data0.condition_data.condition_no = _
      Darea_relation(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Darea_relation(n%).data(0).record.data0.condition_data.condition_no
   Darea_relation(n%).data(0).record.data0.condition_data.condition(i%) = _
     Darea_relation(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'****************************************************
Case angle3_value_
If operat_type = 0 Then
 re.record_data = angle3_value(n%).data(0).record
 re.record_ = angle3_value(n%).record_
 If init_display And angle3_value(n%).record_.display_no > 0 Then
  angle3_value(n%).record_.display_no = 0
 End If
ElseIf operat_type = 1 Then
 angle3_value(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 angle3_value(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   angle3_value(n%).data(0).record.data0.condition_data.condition_no = _
      angle3_value(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To angle3_value(n%).data(0).record.data0.condition_data.condition_no
   angle3_value(n%).data(0).record.data0.condition_data.condition(i%) = _
      angle3_value(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'*******************************************************
Case Dangle_
If operat_type = 0 Then
 re.record_data = Dangle(n%).data(0).record
 re.record_ = Dangle(n%).record_
  If init_display And Dangle(n%).record_.display_no > 0 Then
   Dangle(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 Dangle(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Dangle(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Dangle(n%).data(0).record.data0.condition_data.condition_no = _
      Dangle(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Dangle(n%).data(0).record.data0.condition_data.condition_no
   Dangle(n%).data(0).record.data0.condition_data.condition(i%) = _
     Dangle(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'****************************************************
Case Dline_
If operat_type = 0 Then
 re.record_data = Dline1(n%).data(0).record
 re.record_ = Dline1(n%).record_
  If init_display And Dline1(n%).record_.display_no > 0 Then
   Dline1(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 Dline1(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Dline1(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Dline1(n%).data(0).record.data0.condition_data.condition_no = _
      Dline1(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Dline1(n%).data(0).record.data0.condition_data.condition_no
   Dline1(n%).data(0).record.data0.condition_data.condition(i%) = _
     Dline1(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'********************************************************
Case dpoint_pair_
If operat_type = 0 Then
 re.record_data = Ddpoint_pair(n%).data(0).record
 re.record_ = Ddpoint_pair(n%).record_
   If init_display And Ddpoint_pair(n%).record_.display_no > 0 Then
    Ddpoint_pair(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 Ddpoint_pair(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Ddpoint_pair(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Ddpoint_pair(n%).data(0).record.data0.condition_data.condition_no = _
      Ddpoint_pair(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Ddpoint_pair(n%).data(0).record.data0.condition_data.condition_no
   Ddpoint_pair(n%).data(0).record.data0.condition_data.condition(i%) = _
     Ddpoint_pair(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'******************************************************
Case eline_
If operat_type = 0 Then
 re.record_data = Deline(n%).data(0).record
 re.record_ = Deline(n%).record_
   If init_display And Deline(n%).record_.display_no > 0 Then
    Deline(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 Deline(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Deline(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Deline(n%).data(0).record.data0.condition_data.condition_no = _
      Deline(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Deline(n%).data(0).record.data0.condition_data.condition_no
   Deline(n%).data(0).record.data0.condition_data.condition(i%) = _
     Deline(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'**********************************************************
Case line3_value_
If operat_type = 0 Then
 re.record_data = line3_value(n%).data(0).record
 re.record_ = line3_value(n%).record_
  If init_display And line3_value(n%).record_.display_no > 0 Then
   line3_value(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 line3_value(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 line3_value(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   line3_value(n%).data(0).record.data0.condition_data.condition_no = _
      line3_value(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To line3_value(n%).data(0).record.data0.condition_data.condition_no
   line3_value(n%).data(0).record.data0.condition_data.condition(i%) = _
     line3_value(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'***********************************************************
Case line_value_
If operat_type = 0 Then
 If n% < 0 Then
 re.record_data.data0.condition_data.condition_no = 2
 re.record_data.data0.condition_data.condition(1).ty = line_value_
 re.record_data.data0.condition_data.condition(1).no = -n%
 re.record_data.data0.condition_data.condition(2) = line_value(-n%).data(0).record.data0.condition_for_value_
 re.record_data.data0.theorem_no = 1
 Else
 re.record_data = line_value(n%).data(0).record
 re.record_ = line_value(n%).record_
  If init_display And line_value(n%).record_.display_no > 0 Then
   line_value(n%).record_.display_no = 0
  End If
 End If
ElseIf operat_type = 1 Then
 line_value(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 line_value(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   line_value(n%).data(0).record.data0.condition_data.condition_no = _
      line_value(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To line_value(n%).data(0).record.data0.condition_data.condition_no
   line_value(n%).data(0).record.data0.condition_data.condition(i%) = _
     line_value(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'****************************************************
Case length_of_polygon_
If operat_type = 0 Then
 re.record_data = length_of_polygon(n%).data(0).record
 re.record_ = length_of_polygon(n%).record_
  If init_display And length_of_polygon(n%).record_.display_no > 0 Then
   length_of_polygon(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 length_of_polygon(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 length_of_polygon(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   length_of_polygon(n%).data(0).record.data0.condition_data.condition_no = _
      length_of_polygon(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To length_of_polygon(n%).data(0).record.data0.condition_data.condition_no
   length_of_polygon(n%).data(0).record.data0.condition_data.condition(i%) = _
     length_of_polygon(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'******************
Case two_line_value_
If operat_type = 0 Then
 re.record_data = two_line_value(n%).data(0).record
 re.record_ = two_line_value(n%).record_
  If init_display And two_line_value(n%).record_.display_no > 0 Then
   two_line_value(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 two_line_value(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 two_line_value(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   two_line_value(n%).data(0).record.data0.condition_data.condition_no = _
      two_line_value(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To two_line_value(n%).data(0).record.data0.condition_data.condition_no
   two_line_value(n%).data(0).record.data0.condition_data.condition(i%) = _
     two_line_value(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'******************
Case midpoint_
If operat_type = 0 Then
 re.record_data = Dmid_point(n%).data(0).record
 re.record_ = Dmid_point(n%).record_
   If init_display And Dmid_point(n%).record_.display_no > 0 Then
    Dmid_point(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 Dmid_point(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Dmid_point(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Dmid_point(n%).data(0).record.data0.condition_data.condition_no = _
      Dmid_point(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Dmid_point(n%).data(0).record.data0.condition_data.condition_no
   Dmid_point(n%).data(0).record.data0.condition_data.condition(i%) = _
     Dmid_point(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'************************************************************
Case paral_
If operat_type = 0 Then
 re.record_data = Dparal(n%).data(0).data0.record
 re.record_ = Dparal(n%).record_
   If init_display And Dparal(n%).record_.display_no > 0 Then
    Dparal(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 Dparal(n%).data(0).data0.record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Dparal(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Dparal(n%).data(0).data0.record.data0.condition_data.condition_no = _
      Dparal(n%).data(0).data0.record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Dparal(n%).data(0).data0.record.data0.condition_data.condition_no
   Dparal(n%).data(0).data0.record.data0.condition_data.condition(i%) = _
     Dparal(n%).data(0).data0.record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'********************************************************
Case parallelogram_
If operat_type = 0 Then
 re.record_data = Dparallelogram(n%).data(0).record
 re.record_ = Dparallelogram(n%).record_
  If init_display And Dparallelogram(n%).record_.display_no > 0 Then
   Dparallelogram(n%).record_.display_no = 0
  End If
ElseIf operat_type = 1 Then
 Dparallelogram(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Dparallelogram(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Dparallelogram(n%).data(0).record.data0.condition_data.condition_no = _
      Dparallelogram(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Dparallelogram(n%).data(0).record.data0.condition_data.condition_no
   Dparallelogram(n%).data(0).record.data0.condition_data.condition(i%) = _
     Dparallelogram(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'**************************************************
Case point3_on_line_
If operat_type = 0 Then
 re.record_data = three_point_on_line(n%).data(0).record
 re.record_ = three_point_on_line(n%).record_
   If init_display And three_point_on_line(n%).record_.display_no > 0 Then
    three_point_on_line(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 three_point_on_line(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 three_point_on_line(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   three_point_on_line(n%).data(0).record.data0.condition_data.condition_no = _
      three_point_on_line(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To three_point_on_line(n%).data(0).record.data0.condition_data.condition_no
   three_point_on_line(n%).data(0).record.data0.condition_data.condition(i%) = _
     three_point_on_line(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'******************************************************
Case point4_on_circle_
If operat_type = 0 Then
 re.record_data = four_point_on_circle(n%).data(0).record
 re.record_ = four_point_on_circle(n%).record_
   If init_display And four_point_on_circle(n%).record_.display_no > 0 Then
    four_point_on_circle(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 four_point_on_circle(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 four_point_on_circle(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   four_point_on_circle(n%).data(0).record.data0.condition_data.condition_no = _
      four_point_on_circle(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To four_point_on_circle(n%).data(0).record.data0.condition_data.condition_no
   four_point_on_circle(n%).data(0).record.data0.condition_data.condition(i%) = _
     four_point_on_circle(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'*********************************************************
Case point3_on_circle
If operat_type = 0 Then
 re.record_data = three_point_on_circle(n%).data(0).record
 re.record_ = three_point_on_circle(n%).record_
   If init_display And three_point_on_circle(n%).record_.display_no > 0 Then
    three_point_on_circle(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 three_point_on_circle(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 three_point_on_circle(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   three_point_on_circle(n%).data(0).record.data0.condition_data.condition_no = _
      three_point_on_circle(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To three_point_on_circle(n%).data(0).record.data0.condition_data.condition_no
   three_point_on_circle(n%).data(0).record.data0.condition_data.condition(i%) = _
     three_point_on_circle(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'********************************************************
Case relation_
If operat_type = 0 Then
 re.record_data = Drelation(n%).data(0).record
 re.record_ = Drelation(n%).record_
   If init_display And Drelation(n%).record_.display_no > 0 Then
    Drelation(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 Drelation(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Drelation(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Drelation(n%).data(0).record.data0.condition_data.condition_no = _
      Drelation(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Drelation(n%).data(0).record.data0.condition_data.condition_no
   Drelation(n%).data(0).record.data0.condition_data.condition(i%) = _
     Drelation(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'********************************************************
Case v_relation_
If operat_type = 0 Then
 re.record_data = v_Drelation(n%).data(0).record
 re.record_ = v_Drelation(n%).record_
   If init_display And v_Drelation(n%).record_.display_no > 0 Then
    v_Drelation(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 v_Drelation(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 v_Drelation(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   v_Drelation(n%).data(0).record.data0.condition_data.condition_no = _
      v_Drelation(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To v_Drelation(n%).data(0).record.data0.condition_data.condition_no
   v_Drelation(n%).data(0).record.data0.condition_data.condition(i%) = _
     v_Drelation(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'********************************************************
Case relation_string_
If operat_type = 0 Then
 re.record_data = relation_string(n%).data(0).record
 re.record_ = relation_string(n%).record_
   If init_display And relation_string(n%).record_.display_no > 0 Then
     relation_string(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
  relation_string(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
  relation_string(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   relation_string(n%).data(0).record.data0.condition_data.condition_no = _
      relation_string(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To relation_string(n%).data(0).record.data0.condition_data.condition_no
    relation_string(n%).data(0).record.data0.condition_data.condition(i%) = _
      relation_string(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'********************************************************
Case similar_triangle_
If operat_type = 0 Then
 re.record_data = Dsimilar_triangle(n%).data(0).record
 re.record_ = Dsimilar_triangle(n%).record_
   If init_display And Dsimilar_triangle(n%).record_.display_no > 0 Then
    Dsimilar_triangle(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 Dsimilar_triangle(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Dsimilar_triangle(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Dsimilar_triangle(n%).data(0).record.data0.condition_data.condition_no = _
      Dsimilar_triangle(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Dsimilar_triangle(n%).data(0).record.data0.condition_data.condition_no
   Dsimilar_triangle(n%).data(0).record.data0.condition_data.condition(i%) = _
     Dsimilar_triangle(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'***********************************************
Case total_equal_triangle_
If operat_type = 0 Then
 re.record_data = Dtotal_equal_triangle(n%).data(0).record
 re.record_ = Dtotal_equal_triangle(n%).record_
   If init_display And Dtotal_equal_triangle(n%).record_.display_no > 0 Then
    Dtotal_equal_triangle(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 Dtotal_equal_triangle(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Dtotal_equal_triangle(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Dtotal_equal_triangle(n%).data(0).record.data0.condition_data.condition_no = _
      Dtotal_equal_triangle(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Dtotal_equal_triangle(n%).data(0).record.data0.condition_data.condition_no
   Dtotal_equal_triangle(n%).data(0).record.data0.condition_data.condition(i%) = _
     Dtotal_equal_triangle(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'*************************************************
Case verti_
If operat_type = 0 Then
 re.record_data = Dverti(n%).data(0).record
 re.record_ = Dverti(n%).record_
   If init_display And Dverti(n%).record_.display_no > 0 Then
    Dverti(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 Dverti(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 Dverti(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   Dverti(n%).data(0).record.data0.condition_data.condition_no = _
      Dverti(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To Dverti(n%).data(0).record.data0.condition_data.condition_no
   Dverti(n%).data(0).record.data0.condition_data.condition(i%) = _
     Dverti(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'****************************************************
Case two_order_equation_
If operat_type = 0 Then
 re.record_data = two_order_equation(n%).data(0).record
 re.record_ = two_order_equation(n%).record_
   If init_display And two_order_equation(n%).record_.display_no > 0 Then
    two_order_equation(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
  two_order_equation(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
  two_order_equation(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   two_order_equation(n%).data(0).record.data0.condition_data.condition_no = _
      two_order_equation(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To two_order_equation(n%).data(0).record.data0.condition_data.condition_no
   two_order_equation(n%).data(0).record.data0.condition_data.condition(i%) = _
     two_order_equation(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'****************************************************
Case tri_function_
If operat_type = 0 Then
 re.record_data = tri_function(n%).data(0).record
 re.record_ = tri_function(n%).record_
   If init_display And tri_function(n%).record_.display_no > 0 Then
    tri_function(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 tri_function(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 tri_function(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   tri_function(n%).data(0).record.data0.condition_data.condition_no = _
      tri_function(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To tri_function(n%).data(0).record.data0.condition_data.condition_no
   tri_function(n%).data(0).record.data0.condition_data.condition(i%) = _
     tri_function(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'****************************************************
Case verti_mid_line_
If operat_type = 0 Then
 re.record_data = verti_mid_line(n%).data(0).record
 re.record_ = verti_mid_line(n%).record_
   If init_display And verti_mid_line(n%).record_.display_no > 0 Then
    verti_mid_line(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 verti_mid_line(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 verti_mid_line(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   verti_mid_line(n%).data(0).record.data0.condition_data.condition_no = _
      verti_mid_line(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To verti_mid_line(n%).data(0).record.data0.condition_data.condition_no
   verti_mid_line(n%).data(0).record.data0.condition_data.condition(i%) = _
     verti_mid_line(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'*************************************************************
Case V_line_value_
If operat_type = 0 Then
 re.record_data = V_line_value(n%).data(0).record
 re.record_ = V_line_value(n%).record_
   If init_display And V_line_value(n%).record_.display_no > 0 Then
    V_line_value(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 V_line_value(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 V_line_value(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   V_line_value(n%).data(0).record.data0.condition_data.condition_no = _
      V_line_value(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To V_line_value(n%).data(0).record.data0.condition_data.condition_no
  V_line_value(n%).data(0).record.data0.condition_data.condition(i%) = _
     V_line_value(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'*************************************************************
Case V_two_line_time_value_
If operat_type = 0 Then
 re.record_data = V_two_line_time_value(n%).data(0).record
 re.record_ = V_two_line_time_value(n%).record_
   If init_display And V_two_line_time_value(n%).record_.display_no > 0 Then
    V_two_line_time_value(n%).record_.display_no = 0
   End If
ElseIf operat_type = 1 Then
 V_two_line_time_value(n%).data(0).record.data0 = re.record_data.data0
ElseIf operat_type = 2 Then
 V_two_line_time_value(n%).record_ = re.record_
End If
If remove_no% > 0 Then
   V_two_line_time_value(n%).data(0).record.data0.condition_data.condition_no = _
      V_two_line_time_value(n%).data(0).record.data0.condition_data.condition_no - 1
  For i% = remove_no% To V_two_line_time_value(n%).data(0).record.data0.condition_data.condition_no
   V_two_line_time_value(n%).data(0).record.data0.condition_data.condition(i%) = _
     V_two_line_time_value(n%).data(0).record.data0.condition_data.condition(i% + 1)
  Next i%
End If
'*************************************************************
Case Else
re.record_data.data0.condition_data.condition_no = 0
End Select

End Sub

'**************************************************
'ÒÑÖªÏÔÊ¾ºÅ£¬Ìõ¼þÀàÐÍ£¬½«ÏÔÊ¾ºÅÊäÈë¼ÇÂ¼
'
'**************************************************
'ÒÑÖªÏÔÊ¾ºÅ£¬Ìõ¼þÀàÐÍ£¬½«ÏÔÊ¾ºÅÊäÈë¼ÇÂ¼
'
'******************************************************
Public Sub set_display_no(ByVal ty As Integer, ByVal n%, _
             ByVal dis_n As Integer)
'ÉèÖÃÏÔÊ¾ºÅ
'If wenti_type = 1 Or wenti_type = 3 Then
' If last_conclusion > 0 Then
'  dis_n = dis_n + 2 - last_conclusion
' Else
'  dis_n = dis_n + 1
' End If
'Else
' dis_n = dis_n + 1
'End If
Select Case ty
Case new_point_
new_point(n%).record_.display_no = dis_n
Case item0_
item0(n%).record.display_no = dis_n
Case area_of_element_
 If n% < 0 Then
 area_of_element(-n%).record_.display_no_ = dis_n '+ 1
 Else
 area_of_element(n%).record_.display_no = dis_n
 End If
Case V_line_value_
V_line_value(n%).record_.display_no = dis_n
Case sides_length_of_triangle_
Sides_length_of_triangle(n%).record_.display_no = dis_n
Case sides_length_of_circle_
Sides_length_of_circle(n%).record_.display_no = dis_n
Case area_of_circle_
area_of_circle(n%).record_.display_no = dis_n
Case area_of_fan_
Area_of_fan(n%).record_.display_no = dis_n
Case equation_
equation(n%).record_.display_no = dis_n
Case line3_value_
line3_value(n%).record_.display_no = dis_n
Case tixing_
Dtixing(n%).record_.display_no = dis_n
Case rhombus_
rhombus(n%).record_.display_no = dis_n
Case long_squre_
Dlong_squre(n%).record_.display_no = dis_n
Case equal_side_triangle_
equal_side_triangle(n%).record_.display_no = dis_n
Case equal_side_right_triangle_
equal_side_right_triangle(n%).record_.display_no = dis_n
'Case equal_side_tixing_
'Dequal_side_tixing(n%).record_.display_no = dis_n
Case tangent_line_
tangent_line(n%).record_.display_no = dis_n
Case epolygon_
 If epolygon(n%).record_.display_no = 0 Then
  epolygon(n%).record_.display_no = dis_n '+ 1
 End If
Case general_string_
 general_string(n%).record_.display_no = dis_n ' + 1
Case equal_arc_
 equal_arc(n%).record_.display_no = dis_n '+ 1
'Case equal_area_triangle_
' equal_area_triangle(n%).record_.display_no = dis_n '+ 1
Case area_relation_
 Darea_relation(n%).record_.display_no = dis_n '+ 1
Case angle3_value_
 angle3_value(n%).record_.display_no = dis_n '+ 1
Case Dangle_
  Dangle(n%).record_.display_no = dis_n '+ 1
Case Dline_
  Dline1(n%).record_.display_no = dis_n '+ 1
Case dpoint_pair_
 Ddpoint_pair(n%).record_.display_no = dis_n '+ 1
Case line_value_
 If n% < 0 Then
 line_value(-n%).record_.display_no_ = dis_n '+ 1
 Else
 line_value(n%).record_.display_no = dis_n '+ 1
 End If
Case length_of_polygon_
 If n% < 0 Then
 length_of_polygon(-n%).record_.display_no_ = dis_n '+ 1
 Else
 length_of_polygon(n%).record_.display_no = dis_n '+ 1
 End If
Case paral_
 Dparal(n%).record_.display_no = dis_n '+ 1
Case parallelogram_
 Dparallelogram(n%).record_.display_no = dis_n '+ 1
Case point3_on_line_
  three_point_on_line(n%).record_.display_no = dis_n '+ 1
Case point4_on_circle_
 four_point_on_circle(n%).record_.display_no = dis_n '+ 1
Case midpoint_
 Dmid_point(n%).record_.display_no = dis_n '+ 1
Case relation_
 Drelation(n%).record_.display_no = dis_n '+ 1
Case v_relation_
 v_Drelation(n%).record_.display_no = dis_n '+ 1
Case relation_string_
 relation_string(n%).record_.display_no = dis_n '+ 1
Case verti_
 Dverti(n%).record_.display_no = dis_n '+ 1
Case eline_
 Deline(n%).record_.display_no = dis_n '+ 1
Case verti_mid_line_
 verti_mid_line(n%).record_.display_no = dis_n '+ 1
'Case Two_angle_value_
' Two_angle_value(n%).record_.display_no = dis_n '+ 1
Case two_line_value_
two_line_value(n%).record_.display_no = dis_n '+ 1
Case tri_function_
tri_function(n%).record_.display_no = dis_n
Case total_equal_triangle_
 Dtotal_equal_triangle(n%).record_.display_no = dis_n ' + 1
Case similar_triangle_
 Dsimilar_triangle(n%).record_.display_no = dis_n '+ 1
End Select

End Sub
Public Function set_display_string0(ByVal ty As Integer, ByVal n%, ByVal condition_tree_no%, _
             concl_or_cond As Byte, conclusion_or_inform As Boolean, add_note As Boolean, _
                    is_same_theorem As Byte, ge_or_tree As Byte, dis_ty As Byte, _
                        is_depend As Boolean) As String '½¨Á¢ÏÔÊ¾Óï¾ä
                     
 'no% Ìõ¼þ»ò½áÂÛÔªµÄ¸öÊý'
Dim i%, j%, k%, l%, t_n%, dir%, w_n%
Dim tn(2) As Integer
Dim tpa(2) As String
Dim dis  As Byte
Dim ts As String
Dim tl(3) As Integer
Dim tp(2) As Integer
Dim stri(2) As String
Dim re1 As total_record_type
Dim re2 As total_record_type
Dim note_string
Dim dis_no(7) As Integer
Dim last_dis_no As Integer
Dim no_display_note As Boolean
Dim is_add_condition As Byte
Dim equal_mark As String
Dim brace_mark$
Dim c_data As condition_data_type
display_add_condition(0) = 0 '
If condition_tree_no% > 0 Then
 re2.record_data.data0.condition_data = condition_tree(condition_tree_no%).conditions.data
 re2.record_.conclusion_no = condition_tree(condition_tree_no%).conclusion_no
 re1 = re2
 'µ±Ç°ÏÔÊ¾µÄÊý¾ÝÀàÐÍºÍÐòºÅ
 n% = condition_tree(condition_tree_no%).condition.no
 ty = condition_tree(condition_tree_no%).condition.ty
ElseIf ty > 0 And n% > 0 Then
 Call record_no(ty, n%, re2, False, 0, 0)
  Call record_no(ty, n%, re1, True, 0, 0)
End If
Select Case ty
Case wenti_cond_
   Exit Function
Case relation_string_
   stri(0) = display_string_(relation_string(n%).data(0).relation_string, dis_ty) + "=0"
Case item0_
   stri(0) = set_display_item0(item0(n%).data(0), dis_ty, True, is_depend)
Case new_point_
   stri(0) = new_point(n%).data(0).display_string
  If add_note Then
    stri(0) = stri(0) + "."
   End If
  Call draw_aid_point(n%)
Case V_line_value_ '
     stri(0) = set_display_string_of_V_line_value(V_line_value(n%).data(0), True)
Case area_of_element_
     re1.record_data.data0.theorem_no = 1
      stri(0) = set_area_element_display_string(area_of_element(Abs(n%)).data(0), dis_ty, is_depend)
Case sides_length_of_triangle_
stri(0) = Sides_length_of_triangle(n%).data(0).record.display_string
          '(62, set_display_triangle(Sides_length_of_triangle(n%).data(0).triangle, is_depend, 1, 0) + _
                              "\\3\\" + display_string_(Sides_length_of_triangle(n%).data(0).value, dis_ty)) '+ "~"
Case area_of_circle_
stri(0) = area_of_circle(n%).data(0).record.display_string
     '(area_of_circle(n%).data(0).circ) + _
          "+" + display_string_(area_of_circle(n%).data(0).value, dis_ty) '+ "~"
Case sides_length_of_circle_
're1 = Sides_length_of_circle(n%).data(0).record
stri(0) = Sides_length_of_circle(n%).data(0).record.display_string
'set_display_circle (Sides_length_of_circle(n%).data(0).circ) + _
          "=" + display_string_(Sides_length_of_circle(n%).data(0).value, dis_ty)
Case area_of_fan_
stri(0) = Area_of_fan(n%).data(0).record.display_string
'LoadResString_(1565, "\\1\\" + m_poi(Area_of_fan(n%).data(0).poi(0)).data(0).data0.name + _
                                        m_poi(Area_of_fan(n%).data(0).poi(1)).data(0).data0.name + _
                                        m_poi(Area_of_fan(n%).data(0).poi(2)).data(0).data0.name + _
                              "\\2\\" + display_string_(Area_of_fan(n%).data(0).value, dis_ty)) '+ "~"
    If is_depend Then
     Call set_depend_from_point(Area_of_fan(n%).data(0).poi(0))
     Call set_depend_from_point(Area_of_fan(n%).data(0).poi(1))
     Call set_depend_from_point(Area_of_fan(n%).data(0).poi(2))
    End If
Case point_type_
'If new_point(n%).data(0).record.data1.aid_condition > 0 And _
'    new_point(n%).data(0).record.data0.condition_data.condition_no = 0 Then
'  stri(0) = aid_display_string(new_point(n%).data(0).record.data1.aid_condition).display_string
'   re1 = new_point(n%).data(0).record
'End If
Case two_line_value_
're1 = two_line_value(n%).data(0).record
 stri(0) = two_line_value(n%).data(0).record.display_string
Case tixing_
 stri(0) = Dtixing(n%).data(0).record.display_string
Case rhombus_
 stri(0) = LoadResString_from_inpcond(-10, _
           set_display_polygon4(Dpolygon4(rhombus(n%).data(0).polygon4_no).data(0), 0, is_depend, 1, 0))
Case long_squre_
' re1 = Dlong_squre(n%).data(0).record
 stri(0) = LoadResString_from_inpcond(-13, _
           set_display_polygon4(Dpolygon4(Dlong_squre(n%).data(0).polygon4_no).data(0), 0, is_depend, 1, 0))
Case Squre
' re1 = Dsqure(n%).data(0).record
 stri(0) = LoadResString_from_inpcond(-12, _
           set_display_polygon4(Dpolygon4(Dsqure(n%).data(0).polygon4_no).data(0), 0, is_depend, 1, 0))
Case equation_
 stri(0) = equation(n%).data(0).record.display_string
Case general_string_
  If general_string(n%).data(0).trans_equal_mark = 0 Then
  equal_mark = "="
  ElseIf general_string(n%).data(0).trans_equal_mark = 1 Then
  equal_mark = LoadResString_(1585, "")
  ElseIf general_string(n%).data(0).trans_equal_mark = 2 Then
  equal_mark = LoadResString_(1590, "")
  ElseIf general_string(n%).data(0).trans_equal_mark = 3 Then
  equal_mark = LoadResString_(1595, "") '"£¾"
  ElseIf general_string(n%).data(0).trans_equal_mark = 4 Then
  equal_mark = LoadResString_(1600, "") ' "£¼"
  End If
 If general_string(n%).record_.display_times = 0 Then
  If re1.record_data.data0.condition_data.condition_no > 0 And re1.record_data.data0.condition_data.condition_no < 9 Then
   If re1.record_data.data0.condition_data.condition(1).no > 0 And _
       re1.record_data.data0.condition_data.condition(re1.record_data.data0.condition_data.condition_no).ty = general_string_ And _
         general_string(n%).record_.conclusion_no > 0 Then
     general_string(n%).data(0).trans_para_for_display = _
       time_string(general_string(n%).data(0).trans_para, _
        general_string(re1.record_data.data0.condition_data.condition(re1.record_data.data0.condition_data.condition_no).no).data(0).trans_para_for_display, True, False)
   Else
       general_string(n%).data(0).trans_para_for_display = general_string(n%).data(0).trans_para
   End If
  Else
       general_string(n%).data(0).trans_para_for_display = general_string(n%).data(0).trans_para
  End If
    If Mid$(general_string(n%).data(0).trans_para_for_display, 1, 1) = "-" Then
     stri(0) = set_display_g_string(general_string(n%), False, dis_ty, is_depend)
    Else
     stri(0) = set_display_g_string(general_string(n%), True, dis_ty, is_depend) '·´ºÅ
    End If
      If general_string(n%).display_con_string <> "" Then
       If minus_string(general_string(n%).display_con_string, stri(0), True, False) <> "0" Then
         If stri(0) <> general_string(n%).display_con_string Then
            stri(0) = general_string(n%).display_con_string + "=" + stri(0)
         End If
       End If
      End If
    If stri(0) = "" Then
     stri(0) = general_string(n%).data(0).value
    End If
 Else
   no_display_note = True
   If Mid$(general_string(n%).data(0).trans_para_for_display, 1, 1) <> "-" Then
    If set_display_g_string_with_c_item(n%, True, stri(0), 0, is_depend) = False Then
     Exit Function
    End If
   Else
     If set_display_g_string_with_c_item(n%, False, stri(0), 0, is_depend) = False Then '·´ºÅ
      Exit Function
     End If
   End If
  End If
  '***************
 If re1.record_data.data0.condition_data.condition(1).no > 0 Then
 If general_string(n%).data(0).trans_para_for_display <> "1" And _
     general_string(n%).data(0).trans_para_for_display <> "-1" And _
       general_string(n%).data(0).trans_para_for_display <> "@1" _
      And stri(0) <> "0" Then
     If Mid$(general_string(n%).data(0).trans_para_for_display, 1, 1) = "-" Then
        stri(0) = Mid$(general_string(n%).data(0).trans_para_for_display, 2, _
                  Len(general_string(n%).data(0).trans_para_for_display) - 1) + "*(" + _
           stri(0) + ")"
     Else
        If InStr(2, general_string(n%).data(0).trans_para_for_display, "+", 0) = 0 And _
            InStr(2, general_string(n%).data(0).trans_para_for_display, "-", 0) = 0 And _
             InStr(2, general_string(n%).data(0).trans_para_for_display, "@", 0) = 0 And _
              InStr(2, general_string(n%).data(0).trans_para_for_display, "#", 0) = 0 Then
          If stri(0) = "1" Then 'µ¥Ïî
                stri(0) = general_string(n%).data(0).trans_para_for_display
                GoTo set_display_string0_gennral_string
          Else
           stri(0) = general_string(n%).data(0).trans_para_for_display + "(" + _
             stri(0) + ")"
          End If
        Else '¶àÏî
         stri(0) = "(" + general_string(n%).data(0).trans_para_for_display + ")(" + _
           stri(0) + ")"
        End If
     End If
 End If
 End If
         If general_string(n%).record_.conclusion_no > 0 Then
          If conclusion_data(general_string(n%).record_.conclusion_no - 1).no(0) = n% Then
               If con_general_string(general_string(n%).record_.conclusion_no).data(0).value <> "" Then
                  stri(0) = stri(0) + equal_mark + _
                   display_string_( _
                      con_general_string(general_string(n%).record_.conclusion_no).data(0).value, dis_ty)
               End If
          End If
         End If
 If general_string(n%).record_.conclusion_no > 0 Then
   If general_string(n%).data(0).record.data0.condition_data.condition_no > 8 Or _
           general_string(n%).data(0).record.data0.condition_data.condition_no = 0 Then
    stri(0) = "  " + stri(0) 'Ê×´ÎÏÔÊ¾
   ElseIf general_string(n%).data(0).record.data0.condition_data.condition( _
                  general_string(n%).data(0).record.data0.condition_data.condition_no).ty = general_string_ Then
     If general_string(n%).record_.conclusion_no = _
             general_string(general_string(n%).data(0).record.data0.condition_data.condition( _
                  general_string(n%).data(0).record.data0.condition_data.condition_no).no).record_.conclusion_no Then
set_display_string0_gennral_string:
        stri(0) = equal_mark + stri(0)
      End If
   End If
 End If
 Case tangent_line_
 If is_depend Then
  For i% = 0 To 2
  If tangent_line(n%).data(0).poi(i%) > 0 Then
     Call set_depend_from_point(tangent_line(n%).data(0).poi(i%))
  End If
  Next i%
 End If
 If m_poi(m_lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(0)).data(0).data0.visible > 0 And _
        m_poi(m_lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(1)).data(0).data0.visible > 0 Then
 If tangent_line(n%).data(0).circ(1) = 0 Then
 stri(0) = LoadResString_(1475, "\\1\\" + m_poi(m_lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(0)).data(0).data0.name + _
                                  m_poi(m_lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(1)).data(0).data0.name + _
                         "\\2\\" + m_poi(m_Circ(tangent_line(n%).data(0).circ(0)).data(0).data0.center).data(0).data0.name + _
                                  "(" + m_poi(m_Circ(tangent_line(n%).data(0).circ(0)).data(0).data0.in_point(1)).data(0).data0.name + _
                            ")")
 Else
  stri(0) = LoadResString_(1480, "\\1\\" + m_poi(m_lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(0)).data(0).data0.name + _
                                        m_poi(m_lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(1)).data(0).data0.name + _
                               "\\2\\" + m_poi(m_Circ(tangent_line(n%).data(0).circ(0)).data(0).data0.center).data(0).data0.name + _
      "(" + m_poi(m_Circ(tangent_line(n%).data(0).circ(0)).data(0).data0.in_point(1)).data(0).data0.name + _
       ")" + "\\3\\" + m_poi(m_Circ(tangent_line(n%).data(0).circ(1)).data(0).data0.center).data(0).data0.name + _
      "(" + m_poi(m_Circ(tangent_line(n%).data(0).circ(1)).data(0).data0.in_point(1)).data(0).data0.name + _
       ")")
 End If
 Else
  If tangent_line(n%).data(0).circ(1) = 0 Then
 stri(0) = LoadResString_(1541, "\\1\\" + m_poi(m_Circ(tangent_line(n%).data(0).circ(0)).data(0).data0.center).data(0).data0.name + _
      "(" + m_poi(m_Circ(tangent_line(n%).data(0).circ(0)).data(0).data0.in_point(1)).data(0).data0.name + _
       ")" + "\\2\\" + m_poi(m_lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(0)).data(0).data0.name + _
         m_poi(m_lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(1)).data(0).data0.name)
 Else
  stri(0) = LoadResString_(1545, "\\1\\" + m_poi(m_Circ(tangent_line(n%).data(0).circ(0)).data(0).data0.center).data(0).data0.name + _
      "(" + m_poi(m_Circ(tangent_line(n%).data(0).circ(0)).data(0).data0.in_point(1)).data(0).data0.name + _
       ")" + "\\2\\" + m_poi(m_Circ(tangent_line(n%).data(0).circ(1)).data(0).data0.center).data(0).data0.name + _
      "(" + m_poi(m_Circ(tangent_line(n%).data(0).circ(1)).data(0).data0.in_point(1)).data(0).data0.name + _
       ")" + "\\3\\" + m_poi(m_lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(0)).data(0).data0.name + _
            m_poi(m_lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(1)).data(0).data0.name)
 End If

 End If
 'Call draw_line(Draw_form, m_lin(tangent_line(n%).data(0).line_no).data(0).data0, _
                    condition, 0)
 If m_poi(m_lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(0)).data(0).data0.visible = 0 Then
     Call set_point_visible(m_lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(0), 1, False)
      'Call draw_point(Draw_form, poi(lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(0)), 0, display)
 End If
 If m_poi(m_lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(1)).data(0).data0.visible = 0 Then
     Call set_point_visible(m_lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(1), 1, False)
     ' Call draw_point(Draw_form, poi(lin(tangent_line(n%).data(0).line_no).data(0).data0.poi(1)), 0, display)
 End If
 If m_lin(tangent_line(n%).data(0).line_no).data(0).data0.visible = 0 Then
  Call set_line_visible(tangent_line(n%).data(0).line_no, 1)
 End If
     'Call draw_line(Draw_form, m_lin(tangent_line(n%).data(0).line_no).data(0).data0, _
                     condition, 0)
're1 = tangent_line(n%).data(0).record
Case equal_side_triangle_
stri(0) = LoadResString_from_inpcond(-18, set_display_triangle(equal_side_triangle(n%).data(0).triangle, is_depend, 1, 0))
're1 = equal_side_triangle(n%).data(0).record
Case equal_side_right_triangle_
stri(0) = LoadResString_from_inpcond(-17, set_display_triangle(equal_side_right_triangle(n%).data(0).triangle, is_depend, 1, 0)) _
're1 = equal_side_right_triangle(n%).data(0).record
Case epolygon_
If epolygon(n%).record_.display_times = 0 Or _
     epolygon(n%).record_.display_times = 1 Then
stri(0) = set_display_Epolygon(epolygon(n%).data(0), is_depend, 0, 0)
ElseIf epolygon(n%).record_.display_times = 2 Then
 stri(0) = new_point(epolygon(n%).record_.display_no).data(0).display_string
  no_display_note = True
End If
'      re1 = epolygon(n%).data(0).record
Case equal_arc_
stri(0) = LoadResString_from_inpcond(-24, "\\0\\" + m_poi(arc(equal_arc(n%).data(0).arc(0)).data(0).poi(0)).data(0).data0.name + _
          "\\1\\" + m_poi(arc(equal_arc(n%).data(0).arc(0)).data(0).poi(1)).data(0).data0.name + "\\2\\" + _
           m_poi(arc(equal_arc(n%).data(0).arc(1)).data(0).poi(0)).data(0).data0.name + _
            "\\3|" + m_poi(arc(equal_arc(n%).data(0).arc(1)).data(0).poi(1)).data(0).data0.name)
If is_depend Then
   For i% = 0 To 1
    Call set_depend_from_point(arc(equal_arc(n%).data(0).arc(0)).data(0).poi(i%))
    Call set_depend_from_point(arc(equal_arc(n%).data(0).arc(1)).data(0).poi(i%))
   Next i%
End If
Case area_relation_
 stri(0) = set_area_relation_display_string(Darea_relation(n%).data(0), dis_ty, is_depend)
'**********************************************************
Case angle3_value_
If concl_or_cond > 0 Then
 If con_angle3_value(concl_or_cond - 1).data(1).data0.angle(0) > 0 And _
     con_angle3_value(concl_or_cond - 1).data(1).data0.value <> "" Then
 stri(0) = set_display_three_angle_value(con_angle3_value(concl_or_cond - 1).data(1), _
              True, dis_ty, is_depend)
 Else
 stri(0) = set_display_three_angle_value(con_angle3_value(concl_or_cond - 1).data(0), _
              True, dis_ty, is_depend)
 End If
Else
stri(0) = set_display_three_angle_value(angle3_value(n%).data(0), True, dis_ty, is_depend)
End If
're1 = angle3_value(n%).data(0).record
'**********************************************************
Case Dangle_
 stri(0) = set_angle_display_string(Dangle(n%).data(0).angle) + "."
 ' re1 = Dangle(n%).data(0).record
'************************************************************
Case Dline_
stri(0) = LoadResString_(1605, "\\1\\" + m_poi(Dline1(n%).data(0).poi(0)).data(0).data0.name + _
                                       m_poi(Dline1(n%).data(0).poi(1)).data(0).data0.name)
  Call C_display_picture.set_dot_line( _
      Dline1(n%).data(0).poi(0), Dline1(n%).data(0).poi(1), 0, 0)
'***************************************************
Case dpoint_pair_
If concl_or_cond = 0 Then
 stri(0) = set_display_point_pair(Ddpoint_pair(n%).data(0).data0, Ddpoint_pair(n%).data(0).record, True, is_depend)
Else
 stri(0) = set_display_point_pair(con_dpoint_pair(concl_or_cond - 1).data(1), record_0, True, is_depend)
End If
 '   re1 = Ddpoint_pair(n%).data(0).record
'*******************************************************
Case eline_
If concl_or_cond = 0 Then
stri(0) = set_display_eline(Deline(n%).data(0), True, is_depend)
Else
stri(0) = set_display_eline(con_eline(concl_or_cond - 1).data(1), True, is_depend)
End If
're1 = Deline(n%).data(0).record
Case midpoint_ 'equal_segment_on_line
 stri(0) = set_display_mid_point(Dmid_point(n%), _
    0, True, is_depend)
 're1 = Dmid_point(n%).data(0).record
'***********************************************************
Case mid_point_line_
 stri(0) = m_poi(mid_point_line(n%).data(0).poi(0)).data(0).data0.name + m_poi(mid_point_line(n%).data(0).poi(1)).data(0).data0.name + _
  "+" + m_poi(mid_point_line(n%).data(0).poi(2)).data(0).data0.name + m_poi(mid_point_line(n%).data(0).poi(3)).data(0).data0.name + _
   "= 2*" = m_poi(mid_point_line(n%).data(0).poi(4)).data(0).data0.name + _
     m_poi(mid_point_line(n%).data(0).poi(5)).data(0).data0.name + "."
 If is_depend Then
    For i% = 0 To 5
     Call set_depend_from_point(mid_point_line(n%).data(0).poi(i%))
    Next i%
 End If
Case length_of_polygon_
 stri(0) = set_display_string_for_length_of_polygon(length_of_polygon(n%), 1, is_depend)
   '  re1 = length_of_polygon(n%).data(0).record
Case line_value_
     If n% < 0 Then
      re1.record_data.data0.theorem_no = 1
        stri(0) = set_display_line_value(line_value(-n%), False, dis_ty)
        Call C_display_picture.set_dot_line(line_value(-n%).data(0).data0.poi(0), line_value(-n%).data(0).data0.poi(1), 0, 0)
     Else
      If line_value(n%).data(0).data0.value_ = line_value(n%).data(0).data0.value Then
       stri(0) = set_display_line_value(line_value(n%), True, dis_ty)
      Else
       stri(0) = set_display_line_value(line_value(n%), True, dis_ty)
      End If
        Call C_display_picture.set_dot_line(line_value(n%).data(0).data0.poi(0), line_value(n%).data(0).data0.poi(1), 0, 0)
     End If
     If is_depend Then
       Call set_depend_from_point(line_value(n%).data(0).data0.poi(0))
       Call set_depend_from_point(line_value(n%).data(0).data0.poi(1))
     End If
'**********************************************************
Case line3_value_
   stri(0) = set_display_three_line_value(line3_value(n%).data(0), True, dis_ty, is_depend)
' re1 = line3_value(n%).data(0).record
'*****************************************************
Case paral_
   stri(0) = set_display_paral(Dparal(n%).data(0).data0, True, is_add_condition, is_depend)
 '     re1 = Dparal(n%).data(0).data0.record
'*********************************************
Case parallelogram_
stri(0) = LoadResString_from_inpcond(-11, _
             set_display_polygon4(Dpolygon4(Dparallelogram(n%).data(0).polygon4_no).data(0), 0, is_depend, 1, 0))
 Call C_display_picture.set_dot_line(Dpolygon4(Dparallelogram(n%).data(0).polygon4_no).data(0).poi(0), _
    Dpolygon4(Dparallelogram(n%).data(0).polygon4_no).data(0).poi(1), 0, 0)
 Call C_display_picture.set_dot_line(Dpolygon4(Dparallelogram(n%).data(0).polygon4_no).data(0).poi(1), _
    Dpolygon4(Dparallelogram(n%).data(0).polygon4_no).data(0).poi(2), 0, 0)
 Call C_display_picture.set_dot_line(Dpolygon4(Dparallelogram(n%).data(0).polygon4_no).data(0).poi(2), _
    Dpolygon4(Dparallelogram(n%).data(0).polygon4_no).data(0).poi(3), 0, 0)
 Call C_display_picture.set_dot_line(Dpolygon4(Dparallelogram(n%).data(0).polygon4_no).data(0).poi(3), _
    Dpolygon4(Dparallelogram(n%).data(0).polygon4_no).data(0).poi(0), 0, 0)
  ' re1 = Dparallelogram(n%).data(0).record
 
'***************************************************
Case point3_on_line_
For i% = 0 To 2
 If m_poi(three_point_on_line(n%).data(0).poi(i%)).data(0).data0.name = empty_char Then
   Call set_point_name(three_point_on_line(n%).data(0).poi(i%), _
         next_char(three_point_on_line(n%).data(0).poi(i%), "", 0, 0))
   'Call draw_point(Draw_form, poi(three_point_on_line(n%).data(0).poi(i%)), _
      0, display)
 End If
Next i%
If is_diameter(three_point_on_line(n%).data(0).poi(0), _
 three_point_on_line(n%).data(0).poi(1), _
  three_point_on_line(n%).data(0).poi(2), k%, c_data) Then
 stri(0) = LoadResString_(1640, "\\1\\" + m_poi(three_point_on_line(n%).data(0).poi(0)).data(0).data0.name + _
                                        m_poi(three_point_on_line(n%).data(0).poi(2)).data(0).data0.name + _
                               "\\2\\" + m_poi(three_point_on_line(n%).data(0).poi(1)).data(0).data0.name + _
                                "[down(" + m_poi(m_Circ(k%).data(0).data0.in_point(1)).data(0).data0.name + ")]")
Else
 stri(0) = LoadResString_from_inpcond(24, _
                                "\\0\\" + m_poi(three_point_on_line(n%).data(0).poi(0)).data(0).data0.name + _
                                "\\1\\" + m_poi(three_point_on_line(n%).data(0).poi(1)).data(0).data0.name + _
                                "\\2\\" + m_poi(three_point_on_line(n%).data(0).poi(2)).data(0).data0.name)
End If
'    re1 = three_point_on_line(n%).data(0).record
If is_depend Then
  For i% = 0 To 2
    Call set_depend_from_point(three_point_on_line(n%).data(0).poi(i%))
  Next i%
End If
'****************************************************************
Case point4_on_circle_
For i% = 0 To 3
 If m_poi(four_point_on_circle(n%).data(0).poi(i%)).data(0).data0.name = empty_char Then
 Call set_point_name(four_point_on_circle(n%).data(0).poi(i%), find_new_char)
'   Call draw_point(Draw_form, poi(four_point_on_circle(n%).data(0).poi(i%)), _
     0, display)
 End If
Next i%
 stri(0) = LoadResString_from_inpcond(23, "\\0\\" + m_poi(four_point_on_circle(n%).data(0).poi(0)).data(0).data0.name + _
                               "\\1\\" + m_poi(four_point_on_circle(n%).data(0).poi(1)).data(0).data0.name + _
                               "\\2\\" + m_poi(four_point_on_circle(n%).data(0).poi(2)).data(0).data0.name + _
                               "\\3\\" + m_poi(four_point_on_circle(n%).data(0).poi(3)).data(0).data0.name)
                               ' re1 = four_point_on_circle(n%).data(0).record
'*************************************************************
If is_depend Then
  For i% = 0 To 3
    Call set_depend_from_point(four_point_on_circle(n%).data(0).poi(i%))
  Next i%
End If
Case relation_
For i% = 0 To 3
 If m_poi(Drelation(n%).data(0).data0.poi(i%)).data(0).data0.name = empty_char Then
 Call set_point_name(Drelation(n%).data(0).data0.poi(i%), find_new_char)
  ' Call draw_point(Draw_form, poi(Drelation(n%).data(0).data0.poi(i%)), 0, display)
 End If
Next i%
  stri(0) = set_display_relation(Drelation(n%), 0, True, 1, dis_ty, is_depend) 'concl_or_cond, True, 1,   1)
  Call C_display_picture.set_dot_line(Drelation(n%).data(0).data0.poi(0), Drelation(n%).data(0).data0.poi(1), 0, 0)
  Call C_display_picture.set_dot_line(Drelation(n%).data(0).data0.poi(2), Drelation(n%).data(0).data0.poi(3), 0, 0)
'****************************************************************
Case similar_triangle_
Call direction_1(Dsimilar_triangle(n%).data(0).direction, tn(0), tn(1), tn(2))
stri(0) = set_display_similar_triangle(Dsimilar_triangle(n%).data(0), True, is_depend)
 '      re1 = Dsimilar_triangle(n%).data(0).record
'****************************************************
Case tri_function_
stri(0) = set_display_tri_function(tri_function(n%).data(0), dis_ty, is_depend)
're1 = tri_function(n%).data(0).record
'*************************************************************
Case total_equal_triangle_
Call direction_1(Dtotal_equal_triangle(n%).data(0).direction, tn(0), tn(1), tn(2))
 stri(0) = set_display_total_equal_triangle(Dtotal_equal_triangle(n%).data(0), True, is_depend)
If conclusion_or_inform Then
  Call C_display_wenti.set_m_string("", stri(0), "", "", "", _
        Dtotal_equal_triangle(n%).data(0).record.data0.theorem_no, -1, w_n%, 1) 'C_display_wenti.m_last_input_wenti_no)
    save_statue = 1
Else
  Call C_display_wenti1.set_m_string("", stri(0), "", "", "", _
        Dtotal_equal_triangle(n%).data(0).record.data0.theorem_no, -1, w_n%, 2) 'C_display_wenti.m_last_input_wenti_no)
    save_statue = 1
End If
Case verti_
'If Dverti(n%).data(0).record.data1.aid_condition > 0 And _
'    Dverti(n%).data(0).record.data0.condition_data.condition_no = 0 Then
'stri(0) = aid_display_string(Dverti(n%).data(0).record.data1.aid_condition).display_string
'Call draw_aid_point(Dverti(n%).data(0).record.data1.aid_condition)
'Else
For i% = 0 To 1
 For j% = 0 To 1
 If m_poi(m_lin(Dverti(n%).data(0).line_no(i%)).data(0).data0.poi(j%)).data(0).data0.name = empty_char Then
  Call set_point_name(m_lin(Dverti(n%).data(0).line_no(i%)).data(0).data0.poi(j%), find_new_char)
   'Call draw_point(Draw_form, poi(lin(Dverti(n%).data(0).line_no(i%)).data(0).data0.poi(j%)), 0, display)
 End If
Next j%
Next i%
  stri(0) = set_display_verti(Dverti(n%).data(0), True, is_depend)
   Call C_display_picture.set_dot_line(0, 0, Dverti(n%).data(0).line_no(0), 0)
   Call C_display_picture.set_dot_line(0, 0, Dverti(n%).data(0).line_no(1), 0)
'*********************************************************
'*************************************************
Case verti_mid_line_
stri(0) = LoadResString_(1375, "\\1\\" + _
      m_poi(m_lin(verti_mid_line(n%).data(0).data0.line_no(0)).data(0).data0.poi(0)).data(0).data0.name + _
      m_poi(m_lin(verti_mid_line(n%).data(0).data0.line_no(0)).data(0).data0.poi(1)).data(0).data0.name + _
                               "\\2\\" + _
       m_poi(verti_mid_line(n%).data(0).data0.poi(0)).data(0).data0.name + _
        m_poi(verti_mid_line(n%).data(0).data0.poi(2)).data(0).data0.name)
If is_depend Then
   For i% = 0 To 1
    Call set_depend_from_point(m_lin(verti_mid_line(n%).data(0).data0.line_no(i%)).data(0).data0.poi(0))
    Call set_depend_from_point(m_lin(verti_mid_line(n%).data(0).data0.line_no(i%)).data(0).data0.poi(1))
   Next i%
    For i% = 0 To 2
    Call set_depend_from_point(verti_mid_line(n%).data(0).data0.poi(i%))
    Next i%
End If
're1 = verti_mid_line(n%).data(0).record
'*********************************************************
End Select
     Call record_no(ty, n%, re1, False, 0, 0) '¶Á³öÃ¿¸ö¼ÇÂ¼
If ge_or_tree = 0 Then
 If re1.record_data.data0.theorem_no > 0 Then
  th_chose(re1.record_data.data0.theorem_no).used = 1 '¼ÍÂ¼Ê¹ÓÃµÄ¶¨Àí
 End If
End If
If add_note Then '×÷×¢ÊÍ
If is_same_theorem > 0 Then 'Í¬Àí
   Call record_no(display_string(is_same_theorem).display_record_type, _
      display_string(is_same_theorem).display_record_no, re2, False, 0, 0)
       note_string = LoadResString_(1650, "\\1\\" + CStr(re2.record_.display_no))
ElseIf no_display_note Then
   note_string = ""
Else
If ty = new_point_ Then '¸¨ÖúÏß
  note_string = LoadResString_(1656, "")
Else
 If re1.record_data.data0.condition_data.condition_no = 0 Then 'ÒÑÖªÌõ¼þ
  If re1.record_.conclusion_no = 0 Then
  If re1.record_data.data0.condition_data.condition(8).ty = new_point_ Then
   stri(0) = LoadResString_(1445, "\\1\\" + stri(0)) '"Éè"
  ElseIf ty <> general_string_ And ty <> epolygon_ Then
   note_string = LoadResString_(3955, "\\1\\" + LoadResString_(2155, ""))
  ElseIf ty = general_string_ Then
   If general_string(n%).record_.display_times = 0 Then
    note_string = LoadResString_(3955, "\\1\\" + LoadResString_(2155, ""))
   End If
  Else
   If epolygon(n%).record_.display_times = 0 Then
    note_string = LoadResString_(3955, "\\1\\" + LoadResString_(2155, ""))
   End If
  End If
  End If
 ElseIf re1.record_data.data0.condition_data.condition_no < 200 Then 'ÍÆÀíÌõ¼þ
  If re1.record_data.data0.condition_data.condition_no = 1 And _
        re1.record_data.data0.condition_data.condition(1).ty = wenti_cond_ Then
        note_string = "(" + CStr(re1.record_data.data0.condition_data.condition(1).no) + ")"
  Else
  note_string = "("
   For i% = 1 To re1.record_data.data0.condition_data.condition_no - 1 'Ç°n-1¸ö¼ÇÂ¼
     Call record_no(re1.record_data.data0.condition_data.condition(i%).ty, _
      re1.record_data.data0.condition_data.condition(i%).no, re2, False, 0, 0) '¶Á³öÃ¿¸ö¼ÇÂ¼
       If re2.record_.display_no > 0 Then 'ÍÆÀí¼ÇÂ¼
        For j% = 1 To last_dis_no
         If re2.record_.display_no = dis_no(j%) Then
          GoTo set_display_string0_next0
         End If
        Next j%
         last_dis_no = last_dis_no + 1
         dis_no(last_dis_no) = re2.record_.display_no
         note_string = note_string + _
         CStr(re2.record_.display_no) + _
         LoadResString_(1670, "")
       Else 'Ìõ¼þ¼ÇÂ¼
        For j% = 1 To last_dis_no
         If re2.record_.display_no = -dis_no(j%) Then '±¾ÍÆÀíÖÐÒÑÓÐ¸Ã¼ÇÂ¼
          GoTo set_display_string0_next0
         End If
        Next j%
        last_dis_no = last_dis_no + 1 '±¾ÍÆÀíÖÐÃ»ÓÐ¸Ã¼ÇÂ¼
         dis_no(last_dis_no) = -re2.record_.display_no 'Ìí¼Ó¸Ã¼ÇÂ¼
          Call C_display_wenti.set_m_depend_no(dis_no(last_dis_no))  '¸ÃÒÑÖªÌõ¼þÒÑ±»Ó¦ÓÃ
         note_string = note_string + CStr(-re2.record_.display_no) + _
          LoadResString_(1670, "")
     End If
set_display_string0_next0:
Next i%
   Call record_no(re1.record_data.data0.condition_data.condition(re1.record_data.data0.condition_data.condition_no).ty, _
   re1.record_data.data0.condition_data.condition(re1.record_data.data0.condition_data.condition_no).no, re2, False, 0, 0)
    For j% = 1 To last_dis_no
     If re2.record_.display_no = dis_no(j%) Then
        note_string = Mid$(note_string, 1, Len(note_string) - 1) + ")"
         GoTo set_display_string0_next1
     ElseIf re2.record_.display_no = -dis_no(j%) Then
        note_string = Mid$(note_string, 1, Len(note_string) - 1) + ")"
         GoTo set_display_string0_next1
     End If
      Next j%
     If re2.record_.display_no <> 0 Then
        note_string = note_string + CStr(Abs(re2.record_.display_no)) + ")"
         If re2.record_.display_no < 0 Then
          Call C_display_wenti.set_m_depend_no(Abs(re2.record_.display_no))
         End If
     ElseIf re2.record_data.data0.condition_data.condition_no = 1 And _
        re2.record_data.data0.condition_data.condition(1).ty = new_point_ Then
       note_string = LoadResString_(1655, "")  'note_string +_
     Else
      If note_string = "(" Then
        note_string = ""
      Else
        note_string = note_string + ")"
      End If
     End If
set_display_string0_next1:
End If
End If
End If
End If
set_display_string0_out:
         'stri(0) = stri(0) + "{" + note_string
If ty = general_string_ Then
 If general_string(n%).record_.display_times = 0 And _
     general_string(n%).data(0).combine_two_item(0) > 0 Then
      general_string(n%).record_.display_times = 1
 Call set_display_string0(ty, n%, condition_tree_no%, 0, conclusion_or_inform, add_note, 0, ge_or_tree, dis_ty, is_depend)
 ElseIf general_string(n%).record_.display_times = 1 Then
  general_string(n%).record_.display_times = 0
 End If
ElseIf ty = epolygon_ Then
 If epolygon(n%).record_.display_times = 1 Then
  If epolygon(n%).record_.display_no > 0 Then
   epolygon(n%).record_.display_times = 2
    Call set_display_string0(ty, n%, condition_tree_no%, 0, conclusion_or_inform, add_note, 0, ge_or_tree, dis_ty, is_depend)
     'ÉèÖÃ±ß³¤
  Else
   epolygon(n%).record_.display_times = 0
  End If
 ElseIf epolygon(n%).record_.display_times = 2 Then
   epolygon(n%).record_.display_times = 0
 End If
End If
Else
set_display_string0 = stri(0)
End If
If conclusion_or_inform Then
    If ty = general_string_ Then
     If general_string(n%).record_.conclusion_no > 0 And _
         (general_string(n%).record_.conclusion_ty = 75 Or _
            general_string(n%).record_.conclusion_ty = 76) Then
      If n% = conclusion_data(general_string(n%).record_.conclusion_no - 1).no(0) Then
       stri(1) = Mid$(stri(0), 2, Len(stri(0)) - 1)
       stri(1) = C_display_wenti.m_condition_for_min_max_value(general_string(n%).record_.conclusion_no - 1) + _
                 stri(1)
                 brace_mark$ = ""
      End If
     End If
    End If
    'Call C_display_wenti.set_m_conclusion_or_condition(w_n%, concl_or_cond)
    Call C_display_wenti.set_m_no(ty, n%, w_n%)
    Call C_display_wenti.set_m_string(w_n%, brace_mark$, stri(0), stri(1), stri(2), note_string, _
                   re1.record_data.data0.theorem_no, concl_or_cond, 1)
    Call set_display_no(ty, n%, w_n%)
    Call C_display_wenti.set_m_condition_data(w_n%, ty, n%)
    Call C_display_wenti.set_m_theorem_no(w_n%, re1.record_data.data0.theorem_no)
   save_statue = 1
Else
    set_display_string0 = stri(0)
    'Call C_display_wenti1.set_m_conclusion_or_condition(w_n%, concl_or_cond)
    'Call C_display_wenti.set_m_no(ty, n%, w_n%)
    'Call C_display_wenti1.set_m_string(w_n%, brace_mark$, stri(0), stri(1), stri(2), note_string, _
                    re1.record_data.data0.theorem_no, concl_or_cond, 2)
    'Call C_display_wenti.set_m_condition_data(w_n%, ty, n%)
    'Call set_display_no(ty, n%, wenti_condition_no + w_n%)
   save_statue = 1
End If
display_add_condition(0) = 0
End Function

Public Function LoadResString_(id%, ByVal mod_string As String)
LoadResString_ = LoadResString0_(LoadResString(id% + regist_data.language), mod_string, 0, 0, 0)
End Function
Public Function LoadResString_from_inpcond(inpcond_no%, ByVal mod_string As String) As String
LoadResString_from_inpcond = LoadResString0_(Trim(inpcond(inpcond_no%).inpcond), mod_string, 0, 0, 0, "")
End Function

Public Function read_string_from_string(st%, s1$, brackets1$, brackets2$, p%, S2$, s3$) As String
Dim j%, k%
s1$ = s1$
p% = InStr(st%, s1$, brackets1$, 0) '¶Á³öµÚÒ»¸öÀ¨ºÅ
If p% = 0 Then
  read_string_from_string = ""
  S2$ = s1$
  s3$ = ""
Else
   If brackets1$ = brackets2$ Then
      j% = InStr(p% + Len(brackets1$), s1$, brackets2$, 0)
   Else
   k% = 1
   For j% = p% + 1 To Len(s1$) - Len(brackets1$) + 1
       If Mid$(s1$, j%, Len(brackets1$)) = brackets1$ Then
          j% = j% + 1
          k% = k% + 1
       ElseIf Mid$(s1$, j%, Len(brackets2$)) = brackets2$ Then
           k% = k% - 1
            If k% = 0 Then
             GoTo read_string_from_string_mark1
            End If
           j% = j% + 1
       End If
   Next j%
j% = 0
   End If
read_string_from_string_mark1:
   If j% = 0 Then
     read_string_from_string = ""
     S2$ = s1$
     s3$ = ""
   Else
     read_string_from_string = Mid$(s1$, p%, j% + Len(brackets2$) - p%)
     S2$ = Mid$(s1$, 1, p% - 1)
     If Len(s1$) >= p% + Len(read_string_from_string) Then
     s3$ = Mid$(s1$, p% + Len(read_string_from_string), Len(s1$) - p% - Len(read_string_from_string) + 1)
     End If
   End If
End If
End Function
Public Function LoadResString0_(loadString As String, ByVal mod_string As String, ty As Byte, _
                                     ByVal old_n%, new_n%, Optional ByVal mod_string_ As String = "") As String
                                     'loadString ÊäÈë×Ö·û´®£¬ÐÞ¸ÄµÄ×Ö·û´®£¬old_n%, new_n%ÐÂ¾ÉÎ»ÖÃ£¬mod_string_ÐÞ¸Ä×Ö·û´®
                                     'ty=0ÓÃÓÚ²Ëµ¥Óï¾ä,
Dim mod_string0 As String
Dim id_string(1) As String
Dim new_string(5) As String
Dim i%
mod_string0 = mod_string
LoadResString0_ = loadString
If ty = 0 Then '½öÓÃÓÚ²Ëµ¥Óï¾ä
new_string(0) = read_loadresstring0(loadString, new_string(2), id_string(0)) 'loadStringÔ­×Ö·û´®,new_string(0)µÚÒ»¸ö¹â±êÇ°µÄ×Ö·û´®
                                                            'new_string(2)Ê£ÓàµÄ,id_string(0)µÚÒ»¸ö¹â±ê
If mod_string = "" Then 'ÎÞÐÞ¸Ä
   LoadResString0_ = new_string(0) + new_string(2)
   If id_string(0) <> "" Then
      LoadResString0_ = LoadResString0_(LoadResString0_, "", ty, 0, 0)
   End If
ElseIf mod_string <> "" Then
   If id_string(0) <> "" Then
    new_string(1) = read_modstring0(mod_string0, id_string(0)) 'new_string(1)µÚÒ»¸ö¹â±ê,id_string(0)¹â±êµÄÐòºÅ
    If new_string(2) = "" Then 'Ê£Óà×Ö·ûÎª¿Õ,½áÊø
     LoadResString0_ = new_string(0) + new_string(1) + new_string(2)
    ElseIf Mid$(new_string(2), 1, 2) = "_~" Then
     If new_string(1) = Chr(13) Then
          LoadResString0_ = new_string(0) + Mid$(new_string(2), 2, Len(new_string(2)) - 1)
     Else
          LoadResString0_ = new_string(0) + new_string(1) + id_string(1) + new_string(2)
     End If
    Else
          LoadResString0_ = new_string(0) + new_string(1) + id_string(1) + new_string(2)
    End If
    Do
        new_string(0) = read_loadresstring0(LoadResString0_, new_string(2), id_string(0))
        If new_string(2) <> "" Then
         LoadResString0_ = new_string(0) + new_string(1) + new_string(2)
        End If
    Loop Until new_string(2) = ""
        LoadResString0_ = LoadResString0_(LoadResString0_, mod_string0, ty, 0, 0)
   Else
    LoadResString0_ = loadString
   End If
End If
ElseIf ty = 1 Then 'Éú³ÉÏÔÊ¾Óï¾ä
       If mod_string <> "" Then
         new_string(1) = read_modstring0(mod_string0, id_string(0))
         If id_string(0) <> "" Then
            id_string(1) = from_id_string_to_dis_id_string(id_string(0))
         '¶Á³öÒªÐÞ¸ÄµÄÓï¾ä¶ÔÓ¦id_string×Ö·û´®
         'loadString=new_string(0)+[ id_string]+new_string(2)
            new_string(0) = read_loadresstring0(loadString, new_string(2), id_string(0))
             If new_string(2) = "" Then 'ÐÞ¸Ä¾äÎ²
                LoadResString0_ = new_string(0) + id_string(1) + new_string(1)
                  new_n% = old_n% + 1
             ElseIf Len(new_string(2)) >= 2 Then
               If Mid$(new_string(2), 1, 2) = "_~" Then
                If new_string(1) = "~" Then
                   new_n% = old_n% + 1
                   LoadResString0_ = new_string(0) + Mid$(new_string(2), 2, Len(new_string(2)) - 1)
                Else
                '
                   new_n% = old_n%
                   If mod_string_ <> "" Then
                      Call cut_string_by_string(mod_string_, new_string(2), new_string(3), new_string(4))
                   End If
                      new_string(5) = "[" + next_id_for_string(id_string(0)) + "]"
                      new_string(2) = next_id_for_string(new_string(2))
                   LoadResString0_ = new_string(0) + id_string(1) + _
                                                      new_string(1) + new_string(5) + _
                                                      next_id_for_string(new_string(2))
                   mod_string_ = new_string(3) + new_string(5) + new_string(2) + new_string(4)
                    new_n% = old_n% + 1
                End If
               ElseIf Mid$(new_string(2), 1, 1) = "_" Then
                   new_n% = old_n% + 1
                   LoadResString0_ = new_string(0) + id_string(1) + _
                                                      new_string(1) + Mid$(new_string(2), 2, Len(new_string(2)) - 1)
               Else
                  new_n% = old_n% + 1
                  LoadResString0_ = new_string(0) + id_string(1) + _
                                                       new_string(1) + new_string(2)
               End If
             ElseIf Len(new_string(2)) = 1 Then
               If Mid$(new_string(2), 1, 1) = "_" Then
                   new_n% = old_n% + 1
                   LoadResString0_ = new_string(0) + id_string(1) + _
                                                      new_string(1) + Mid$(new_string(2), 2, Len(new_string(2)) - 1)
               Else
                  new_n% = old_n% + 1
                  LoadResString0_ = new_string(0) + id_string(1) + _
                                                     new_string(1) + new_string(2)
               End If
             Else
                  LoadResString0_ = new_string(0) + id_string(1) _
                                                     + new_string(1) + new_string(2)
             End If
       Do
        new_string(0) = read_loadresstring0(LoadResString0_, new_string(2), id_string(0))
        If id_string(0) <> "" Then
         LoadResString0_ = new_string(0) + id_string(1) + new_string(1) + new_string(2)
        End If
       Loop Until id_string(0) = ""
        LoadResString0_ = LoadResString0_(LoadResString0_, mod_string0, ty, 0, 0)
     Else
      LoadResString0_ = loadString
    End If
   Else
          LoadResString0_ = loadString
   End If
End If
End Function

Private Function read_loadresstring0(ByVal s1 As String, S2 As String, id_string As String) As String '
's1 Ô­×Ö·û
's2 ±êºÅ¶ÔÓ¦µÄ×Ö·û
'id_string ±êºÅ
'read_loadresstring0 ·µ»ØÊ£ÓàµÄ×Ö·û
'ÔÚÖ¸¶¨id_string ±êºÅ£¬½«×Ö·û´®·ÖÎªÁ½¶Î
Dim i%
id_string = Trim(id_string)
s1 = Trim(s1)
If id_string <> "" Then
   i% = InStr(1, s1, id_string, 0) 'id_stringµÄÎ»ÖÃ
   If i% = 0 Then 'ÎÞid_string
    id_string = ""
     read_loadresstring0 = s1
      S2 = ""
   Else
     read_loadresstring0 = Mid$(s1, 1, i% - 1) 'Ç°¶Î
     If read_loadresstring0 <> "" Then
         If Mid$(read_loadresstring0, Len(read_loadresstring0), 1) = "[" Then
             read_loadresstring0 = Mid$(read_loadresstring0, 1, Len(read_loadresstring0) - 1)
         End If
     End If
     S2 = Mid$(s1, i% + 5, Len(s1) - i% - 4) 'ºó¶Î
     If S2 <> "" Then
         If Mid$(S2, 1, 1) = "]" Then
             S2 = Mid$(S2, 2, Len(S2) - 1)
         End If
     End If
   End If
Else
 id_string = read_string_from_string(1, s1, "\\", "\\", 0, read_loadresstring0, S2)
 If id_string <> "" Then
    If read_loadresstring0 <> "" Then
       If Mid$(read_loadresstring0, Len(read_loadresstring0), 1) = "[" Then
          read_loadresstring0 = Mid$(read_loadresstring0, 1, Len(read_loadresstring0) - 1)
       End If
    End If
    If S2 <> "" Then
       If Mid$(S2, 1, 1) = "]" Then
          S2 = Mid$(S2, 2, Len(S2) - 1)
       End If
    End If
    If S2 <> "" Then
      If Mid$(S2, 1, 1) = "_" Then
          S2 = Mid$(S2, 2, Len(S2) - 1)
      End If
    End If
 End If
End If
End Function
Private Function read_modstring0(Mods As String, id_string As String) As String '
Dim i%, j%
Dim ts$
If id_string = "" Then
   If Mods <> "" Then
    '¶Á³öÐÞ¸ÄÓï¾äµÄ±êºÅ,
    id_string = read_string_from_string(1, Mods, "\\", "\\", 0, "", ts$)
    If id_string <> "" And ts <> "" Then
    j% = InStr(1, ts$, "\\", 0)
    If j% > 0 Then
     read_modstring0 = Mid$(ts$, 1, j% - 1)
     Mods = Mid$(ts, j% + 2, Len(ts$) - j% - 1)
    Else
     read_modstring0 = ts$
     Mods = ""
    End If
   Else
    Mods = Mid$(Mods, 1 + Len(id_string), Len(Mods) - Len(id_string))
    id_string = ""
    read_modstring0 = ""
   End If
   Else
   read_modstring0 = ""
   End If
Else
i% = InStr(1, Mods, id_string, 0)
If i% > 0 Then
   j% = InStr(i% + 5, Mods, "\\", 0)
   If j% > 0 Then
      read_modstring0 = Mid$(Mods, i% + 5, j - i% - 5) 'Ç°¶Î
      Mods = Mid$(Mods, 1, i% - 1) + Mid$(Mods, j%, Len(Mods) - j% + 1) 'ºó¶Î
   Else
      read_modstring0 = Mid$(Mods, i% + 5, Len(Mods) - i% - 4)
      Mods = Mid$(Mods, 1, i% - 1)
   End If
Else
      read_modstring0 = ""
End If
End If
End Function
Public Function read_id_string(st%, ts$, p%, n%) As String
'´Ó¶Á³öµÚp%¶Î ³¤n%
Dim i%, j%
read_id_string = read_string_from_string(st%, ts$, "\\", "\\", p%, "", "")
 If read_id_string <> "" Then
  n% = val(Mid$(read_id_string, 3, Len(read_id_string) - 4))
 Else
  n% = -1
 End If
End Function
Public Function replace_string_by_string(s1$, old_s$, new_string As String) As Boolean
Dim i%
Dim ty As Boolean
Dim S2$
Dim s3$
If Trim(old_s$) <> Trim(new_string) Then
Do
ty = cut_string_by_string(s1$, old_s$, S2$, s3$)
If ty Then
  s1$ = S2$ + new_string + s3$
  replace_string_by_string = True
End If
Loop Until ty = False
End If
End Function
Public Function cut_string_by_string(s1$, old_s$, S2$, s3$) As Boolean
Dim i%
i% = InStr(1, s1, old_s$, 0)
If i% = 0 Then
   cut_string_by_string = False
Else
   cut_string_by_string = True
   S2$ = Mid$(s1$, 1, i% - 1) '
   If Len(s1$) >= i% + Len(old_s$) Then
   s3$ = Mid$(s1$, i% + Len(old_s$), Len(s1$) - i% - Len(old_s$) + 1)
   Else
   s3$ = ""
   End If
End If
End Function

Public Function string_delete_string(s1 As String, S2 As String) As Boolean
Dim k%
k% = InStr(1, s1, S2, 0)
Do
If k% > 0 Then
 s1 = Mid$(s1, 1, k% - 1) + Mid$(s1, k% + Len(S2), Len(s1) - k% - Len(S2) + 1)
End If
k% = InStr(1, s1, S2, 0)
Loop Until k% = 0
End Function
Public Function simple_display_string(in_s As String) As String
Dim ts As String
Dim s1 As String
Dim S2 As String
simple_display_string = in_s
ts = read_string_from_string(1, simple_display_string, "[", "]", 0, s1, S2)
If ts <> "" Then
Do
   ts = Mid$(ts, 2, Len(ts) - 2)
   Call string_delete_string(ts, "down")
   Call string_delete_string(ts, "up")
   simple_display_string = s1 + ts + S2
  ts = read_string_from_string(1, simple_display_string, "[", "]", 0, s1, S2)
Loop Until ts = ""
End If
Call string_delete_string(simple_display_string, "!")
Call string_delete_string(simple_display_string, "~")
End Function
Public Function next_id_for_string(ByVal inS As String) As String
Dim ts(2) As String
Dim id_string As String
id_string = read_string_from_string(1, inS, "\\", "\\", 0, ts(0), ts(1))
If id_string = "" Then
    next_id_for_string = inS$
Else
   ts(2) = Mid$(id_string, 3, Len(id_string) - 4)
   next_id_for_string = ts(0) & "\\" & Trim(str(val(ts(2) + 1))) & "\\" & next_id_for_string(ts(1))
End If
End Function

Private Function from_id_string_to_dis_id_string(ByVal IdString As String) As String
 from_id_string_to_dis_id_string = "{{" + Mid$(IdString, 3, Len(IdString) - 4) + "}}"
End Function


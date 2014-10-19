Attribute VB_Name = "c_list"
Option Explicit

Public Sub set_gate1(volume%) '显示第一册目录的过程函数“set_gate1(1)”
 If volume% = 1 Then          '条件判断语句
clist_form.List6.AddItem "                                                "
clist_form.List6.AddItem "                                                "
clist_form.List6.AddItem "                  平面几何    第一册"
clist_form.List6.AddItem "                                                "
clist_form.List6.AddItem "                                                "
clist_form.List6.AddItem "                                                "
clist_form.List6.AddItem "  第一章   线段 角"  '在List6列表框中显示的各章节目录
clist_form.List6.AddItem "                                                "
clist_form.List6.AddItem "      一     直线、射线、线段"
clist_form.List6.AddItem "         1.1     直线    "
clist_form.List6.AddItem "         1.2    射线、线段"
clist_form.List6.AddItem "         1.3   线段的比较和画法"
clist_form.List6.AddItem "       读一读 长度单位"
clist_form.List6.AddItem "      二     角"
clist_form.List6.AddItem "         1.4     角"
clist_form.List6.AddItem "         1.5     角的比较"
clist_form.List6.AddItem "         1.6     角的度量"
clist_form.List6.AddItem "       读一读  角的度量和六十进制"
clist_form.List6.AddItem "         1.7     角的画法"
clist_form.List6.AddItem "       小结与复习"
clist_form.List6.AddItem "       复习题一"
clist_form.List6.AddItem "       自我测验一"
clist_form.List6.AddItem "                                                "
clist_form.List6.AddItem "   第二章   相交线、平行线"
clist_form.List6.AddItem "                                                "
clist_form.List6.AddItem "      一     相交线、垂线"
clist_form.List6.AddItem "         2.1     相交线、对顶角"
clist_form.List6.AddItem "         2.2     垂线"
clist_form.List6.AddItem "         2.3     同位角、内错角、同旁内角"
clist_form.List6.AddItem "      二     平行线"
clist_form.List6.AddItem "         2.4     平行线及平行公理 "
clist_form.List6.AddItem "       读一读  观察与实验"
clist_form.List6.AddItem "         2.5     平行线的判定"
clist_form.List6.AddItem "         2.6     平行线的性质"
clist_form.List6.AddItem "         2.7     空间里的平行关系"
clist_form.List6.AddItem "      三     命题 定理 证明"
clist_form.List6.AddItem "         2.9     命题 "
clist_form.List6.AddItem "         2.10     定理与证明 "
clist_form.List6.AddItem "       读一读 推理"
clist_form.List6.AddItem "       小结与复习"
clist_form.List6.AddItem "       复习题二"
clist_form.List6.AddItem "       自我测验二"
clist_form.List6.AddItem "       读一读  有关几何的一些历史"
clist_form.List6.AddItem "           附录  部分习题答案"
 End If
End Sub
Public Sub set_gate2(volume%) '显示第二册目录的过程函数“set_gate1(1)”
 If volume% = 2 Then          '条件判断语句
  clist_form.List6.AddItem "                                                "
  clist_form.List6.AddItem "                                                "
  clist_form.List6.AddItem "                 平面几何 第二册"
  clist_form.List6.AddItem "                                                "
  clist_form.List6.AddItem "                                                "
  clist_form.List6.AddItem "   第三章  三角形"
  clist_form.List6.AddItem "                                                "
  clist_form.List6.AddItem "        一   三角形"
  clist_form.List6.AddItem "              3.1  关于三角形的一些概念"
  clist_form.List6.AddItem "              3.2  三角形三条边的关系"
  clist_form.List6.AddItem "              3.3  三角形的内角和"
  clist_form.List6.AddItem "        二    全等三角形"
  clist_form.List6.AddItem "              3.4  全等三角形"
  clist_form.List6.AddItem "           读一读 全等变换"
  clist_form.List6.AddItem "              3.5  三角形全等的判定(一)"
  clist_form.List6.AddItem "              3.6  三角形全等的判定(二)"
  clist_form.List6.AddItem "              3.7  三角形全等的判定(三)"
  clist_form.List6.AddItem "              3.8  直角三角形全等的判定"
  clist_form.List6.AddItem "              3.9  角的平分线"
  clist_form.List6.AddItem "        三  尺规作图"
  clist_form.List6.AddItem "              3.10   基本作图"
  clist_form.List6.AddItem "              3.11  作图题举例"
  clist_form.List6.AddItem "           读一读 关于三等分角的问题"
  clist_form.List6.AddItem "        四   等腰三角形"
  clist_form.List6.AddItem "              3.12 等腰三角形的性质 "
  clist_form.List6.AddItem "              3.13 等腰三角形的判定"
  clist_form.List6.AddItem "           读一读  三角形中边与角之间的不等关系"
  clist_form.List6.AddItem "              3.14  线段的垂直平分线"
  clist_form.List6.AddItem "              3.15  轴对称和轴对称图形"
  clist_form.List6.AddItem "        五    勾股定理"
  clist_form.List6.AddItem "              3.16  勾股定理"
  clist_form.List6.AddItem "              3.17  勾股定理的逆定理"
  clist_form.List6.AddItem "           读一读 勾股定理的证明 "
  clist_form.List6.AddItem "            小结与复习"
  clist_form.List6.AddItem "            复习题三"
  clist_form.List6.AddItem "            自我测验三"
  clist_form.List6.AddItem "                                                "
  clist_form.List6.AddItem "    第四章 四边形"
  clist_form.List6.AddItem "                                                "
  clist_form.List6.AddItem "       一  四边形"
  clist_form.List6.AddItem "            4.1 四边形  "
  clist_form.List6.AddItem "            4.2 多边形的内角和"
  clist_form.List6.AddItem "          读一读 巧用材料"
  clist_form.List6.AddItem "       二   平行四边形"
  clist_form.List6.AddItem "            4.3 平行四边形及其性质"
  clist_form.List6.AddItem "            4.4 平行四边形的判定"
  clist_form.List6.AddItem "            4.5 矩形、菱形"
  clist_form.List6.AddItem "            4.6 正方形"
  clist_form.List6.AddItem "          读一读  完美的正方形"
  clist_form.List6.AddItem "            4.7 中心对称和中心对称图形"
  clist_form.List6.AddItem "        三  梯形"
  clist_form.List6.AddItem "            4.8  梯形"
  clist_form.List6.AddItem "            4.9  平行线等分线段定理"
  clist_form.List6.AddItem "            4.10 三角形、梯形的中位线    "
  clist_form.List6.AddItem "            4.11 不规则多边形的面积"
  clist_form.List6.AddItem "             小结与复习"
  clist_form.List6.AddItem "             复习题四"
  clist_form.List6.AddItem "             自我测验四"
  clist_form.List6.AddItem "                                                "
  clist_form.List6.AddItem "     第五章  相似形"
  clist_form.List6.AddItem "                                                "
  clist_form.List6.AddItem "         一  比例线段"
  clist_form.List6.AddItem "              5.1  比例线段"
  clist_form.List6.AddItem "            读一读  黄金分割"
  clist_form.List6.AddItem "              5.2  平行线分线段成比例定理"
  clist_form.List6.AddItem "         二  相似三角形"
  clist_form.List6.AddItem "              5.3 相似三角形"
  clist_form.List6.AddItem "              5.4 相似三角形的判定"
  clist_form.List6.AddItem "              5.5 相似三角形的性质  "
  clist_form.List6.AddItem "              5.6  相似多边形"
  clist_form.List6.AddItem "             读一读 位似变换"
  clist_form.List6.AddItem "               小结与复习"
  clist_form.List6.AddItem "               复习题五"
  clist_form.List6.AddItem "               自我测验五"
  clist_form.List6.AddItem "         附录 部分习题答案或提示"
   End If
 End Sub
Public Sub set_gate3(volume%)  '显示第三册目录的过程函数“set_gate1(1)”
 If volume% = 3 Then           '条件判断语句
 clist_form.List6.AddItem "                                                "
 clist_form.List6.AddItem "                                                "
 clist_form.List6.AddItem "                    平面几何 第三册"
 clist_form.List6.AddItem "                                                "
 clist_form.List6.AddItem "                                                "
 clist_form.List6.AddItem "  第六章  解直角三角形"
 clist_form.List6.AddItem "                                                "
 clist_form.List6.AddItem "      一  锐角三角形"
 clist_form.List6.AddItem "           6.1  正弦和余弦  "
 clist_form.List6.AddItem "           6.2  正切和余切"
 clist_form.List6.AddItem "      二  解直角三角形 "
 clist_form.List6.AddItem "           6.3  解直角三角形"
 clist_form.List6.AddItem "           6.4  应用举例"
 clist_form.List6.AddItem "        读一读 中国古代有关三角的一些研究"
 clist_form.List6.AddItem "           6.5  实习作业"
 clist_form.List6.AddItem "        小节  与复习 "
 clist_form.List6.AddItem "        复习题六"
 clist_form.List6.AddItem "        自我测验六"
 clist_form.List6.AddItem "                                                "
 clist_form.List6.AddItem "  第七章  圆"
 clist_form.List6.AddItem "                                                "
 clist_form.List6.AddItem "      一  圆的有关性质"
 clist_form.List6.AddItem "           7.1   圆"
 clist_form.List6.AddItem "           7.2   过三点的圆"
 clist_form.List6.AddItem "         读一读 轨迹在作图中的应用"
 clist_form.List6.AddItem "           7.3  垂直于弦的直径"
 clist_form.List6.AddItem "           7.4  圆心角、弧、弦、弦心距之间的关系"
 clist_form.List6.AddItem "           7.5  圆周角"
 clist_form.List6.AddItem "           7.6  圆的内接四边形"
 clist_form.List6.AddItem "      二  直线和圆的位置关系"
 clist_form.List6.AddItem "           7.7  直线和圆的位置关系"
 clist_form.List6.AddItem "           7.8  切线的判定和性质"
 clist_form.List6.AddItem "        读一读 为什么车轮做成圆的？"
 clist_form.List6.AddItem "           7.9  三角形的内切圆"
 clist_form.List6.AddItem "           7.10 切线长定理"
 clist_form.List6.AddItem "           7.11 弦切角"
 clist_form.List6.AddItem "           7.12 和圆有关的比例线段"
 clist_form.List6.AddItem "      三  圆和圆的位置关系"
 clist_form.List6.AddItem "           7.13 圆和圆的位置关系"
 clist_form.List6.AddItem "           7.14 两圆的公切线"
 clist_form.List6.AddItem "           7.15 相切在作图中的应用"
 clist_form.List6.AddItem "      四  正多边形和圆"
 clist_form.List6.AddItem "           7.16 正多边形和圆"
 clist_form.List6.AddItem "           7.17  正多边形的有关计算"
 clist_form.List6.AddItem "         读一读 旋转对称"
 clist_form.List6.AddItem "           7.18  画正多边形"
 clist_form.List6.AddItem "           7.19  圆周长、弧长"
 clist_form.List6.AddItem "         读一读 圆周率"
 clist_form.List6.AddItem "           7.20   圆、扇形、弓形的面积"
 clist_form.List6.AddItem "         读一读 等周问题"
 clist_form.List6.AddItem "           7.21   圆柱和圆锥的侧面展开图"
 clist_form.List6.AddItem "       小节与复习"
 clist_form.List6.AddItem "       复习题七"
 clist_form.List6.AddItem "       自我测验七"
 clist_form.List6.AddItem "           附录 部分习题答案或提示"
   End If
 End Sub

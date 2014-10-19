Attribute VB_Name = "draw_form"
Option Explicit

Public Sub draw_hu(ob As Object, ByVal cx As Single, ByVal cy As Single, ByVal startx As Single, _
       ByVal starty As Single, ByVal endx As Single, ByVal endy As Single, ByVal color%)
Dim r As Single
'Dim fm As Object
'Set fm = Form1
Const PI = 3.14159265358979
    r = Sqr((startx - cx) ^ 2 + (starty - cy) ^ 2)
If starty < cy And startx > cx And endx > cx And endy < cy Then '第一象限到第一象限
         ob.Circle (cx, cy), r, QBColor(color%), Atn((cy - starty) / (startx - cx)), Atn((cy - endy) / (endx - cx))
ElseIf startx > cx And starty < cy And endx < cx And endy < cy Then '第一象限到第二象限
         ob.Circle (cx, cy), r, QBColor(color%), Atn((cy - starty) / (startx - cx)), PI / 2 + Atn((cy - endy) / (cx - endx))
ElseIf startx > cx And starty < cy And endx < cx And endy > cy Then '第一象限到第三象限
         ob.Circle (cx, cy), r, QBColor(color%), Atn((cy - starty) / (startx - cx)), PI + Atn((endy - cy) / (cx - endx))
ElseIf startx > cx And starty < cy And endx > cx And endy > cy Then '第一象限到第四象限
         ob.Circle (cx, cy), r, QBColor(color%), Atn((cy - starty) / (startx - cx)), 2 * PI - Atn((endy - cy) / (endx - cx))
ElseIf startx < cx And starty < cy And endx > cx And endy < cy Then '第二象限到第一象限
         ob.Circle (cx, cy), r, QBColor(color%), PI / 2 + Atn((cy - starty) / (cx - startx)), Atn((cy - endy) / (endx - cx))
ElseIf startx < cx And starty < cy And endx < cx And endy < cy Then '第二象限到第二象限
         ob.Circle (cx, cy), r, QBColor(color%), PI / 2 + Atn((cy - starty) / (cx - startx)), PI / 2 + Atn((cy - endy) / (cx - endx))
ElseIf startx < cx And starty < cy And endx < cx And endy > cy Then '第二象限到第三象限
         ob.Circle (cx, cy), r, QBColor(color%), PI / 2 + Atn((cy - starty) / (cx - startx)), PI + Atn((endy - cy) / (cx - endx))
ElseIf startx < cx And starty < cy And endx > cx And endy > cy Then '第二象限到第四象限
         ob.Circle (cx, cy), r, QBColor(color%), PI / 2 + Atn((cy - starty) / (cx - startx)), 2 * PI - Atn((endy - cy) / (endx - cx))
ElseIf starty > cy And startx < cx And endx > cx And endy < cy Then '第三象限到第一象限
         ob.Circle (cx, cy), r, QBColor(color%), PI + Atn((starty - cy) / (cx - startx)), Atn((cy - endy) / (endx - cx))
ElseIf startx < cx And starty > cy And endx < cx And endy < cy Then '第三象限到第二象限
         ob.Circle (cx, cy), r, QBColor(color%), PI + Atn((starty - cy) / (cx - startx)), PI / 2 + Atn((endy - cy) / (endx - cx))
ElseIf startx < cx And starty > cy And endx < cx And endy > cy Then '第三象限到第三象限
         ob.Circle (cx, cy), r, QBColor(color%), PI + Atn((starty - cy) / (cx - startx)), PI + Atn((endy - cy) / (cx - endx))
ElseIf startx < cx And starty > cy And endx > cx And endy > cy Then '第三象限到第四象限
         ob.Circle (cx, cy), r, QBColor(color%), PI + Atn((starty - cy) / (cx - startx)), 2 * PI - Atn((endy - cy) / (endx - cx))
ElseIf starty > cy And startx > cx And endy < cy And endx > cx Then '第四象限到第一象限
         ob.Circle (cx, cy), r, QBColor(color%), 2 * PI - Atn((starty - cy) / (startx - cx)), Atn((cy - endy) / (endx - cx))
ElseIf startx > cx And starty > cy And endx < cx And endy < cy Then '第四象限到第二象限
         ob.Circle (cx, cy), r, QBColor(color%), 2 * PI - Atn((starty - cy) / (startx - cx)), PI / 2 + Atn((cy - endy) / (cx - endx))
ElseIf startx > cx And starty > cy And endx < cx And endy > cy Then '第四象限到第三象限
         ob.Circle (cx, cy), r, QBColor(color%), 2 * PI - Atn((starty - cy) / (startx - cx)), PI + Atn((endy - cy) / (cx - endx))
ElseIf startx > cx And starty > cy And endx > cx And endy > cy Then '第四象限到第四象限
         ob.Circle (cx, cy), r, QBColor(color%), 2 * PI - Atn((starty - cy) / (startx - cx)), 2 * PI - Atn((endy - cy) / (endx - cx))
End If
Exit Sub
End Sub


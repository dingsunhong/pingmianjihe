Attribute VB_Name = "set_reduce"
Option Explicit
Dim connect_point() As Integer
Type branch_data_type
 branch_no As Integer
 sbus_no As Integer
 sbus_to As Integer
 chain_from As Integer
 chain_to As Integer
End Type
Global branch_data() As branch_data_type
Type conclusion_point_type
poi(20) As Integer
End Type
Global conclusion_point(3) As conclusion_point_type
Dim last_connect_point As Integer
Public Sub change_depend_point_for_line(ByVal l%, ByVal p%, n%)
Dim i%, p1%, p2%, p3%
If m_lin(l%).data(0).parent.element((n% + 1) Mod 2).ty = point_ Then
    p1% = m_lin(l%).data(0).parent.element((n% + 1) Mod 2).no
End If
p2% = p%
For i% = 1 To m_lin(l%).data(0).data0.in_point(0)
     If m_poi(m_lin(l%).data(0).data0.in_point(i%)).data(0).parent.element(1).no > 0 Then
     If m_lin(l%).data(0).data0.in_point(i%) <> p1% And _
          m_lin(l%).data(0).data0.in_point(i%) <> p2% Then
         If p3% = 0 Then
            p3% = m_lin(l%).data(0).data0.in_point(i%)
         End If
         If m_lin(l%).data(0).data0.in_point(i%) > p% Then
             If p2% = p% Then
                p2% = m_lin(l%).data(0).data0.in_point(i%)
             ElseIf p2% > m_lin(l%).data(0).data0.in_point(i%) Then
                p2% = m_lin(l%).data(0).data0.in_point(i%)
             End If
         End If
     End If
     End If
Next i%
    If p2% = p% Then
     p2% = p3%
    End If
    m_poi(p%).data(0).parent.element(1).ty = line_
    m_poi(p%).data(0).parent.element(1).no = l%
    m_lin(l%).data(0).parent.element(n%).ty = point_
    m_lin(l%).data(0).parent.element(n%).no = p2%
    If m_poi(p2%).data(0).parent.element(0).ty = line_ And _
         m_poi(p2%).data(0).parent.element(0).no = l% Then
         m_poi(p2%).data(0).parent.element(0).ty = _
           m_poi(p2%).data(0).parent.element(1).ty
         m_poi(p2%).data(0).parent.element(0).no = _
           m_poi(p2%).data(0).parent.element(1).no
         m_poi(p2%).data(0).parent.element(1).ty = 0
         m_poi(p2%).data(0).parent.element(1).no = 0
    ElseIf m_poi(p2%).data(0).parent.element(1).ty = line_ And _
         m_poi(p2%).data(0).parent.element(1).no = l% Then
         m_poi(p2%).data(0).parent.element(1).ty = 0
         m_poi(p2%).data(0).parent.element(1).no = 0
    End If
End Sub
Public Function is_point_depend_line(ByVal p%, ByVal l%) As Byte
If m_poi(p%).data(0).parent.element(0).ty = line_ And _
    m_poi(p%).data(0).parent.element(0).no = l% Then
     is_point_depend_line = True
ElseIf m_poi(p%).data(0).parent.element(1).ty = line_ And _
    m_poi(p%).data(0).parent.element(1).no = l% Then
     is_point_depend_line = True
End If
End Function

Public Function is_point_in_item(ByVal it_no%, ByVal p%) As Boolean
 If is_point_in_element(item0(it_no%).data(0).poi(0), item0(it_no%).data(0).poi(1), p%) Then
     is_point_in_item = True
 ElseIf is_point_in_element(item0(it_no%).data(0).poi(2), item0(it_no%).data(0).poi(3), p%) Then
     is_point_in_item = True
 End If
End Function
Public Function is_point_in_element(ByVal p1%, ByVal p2%, ByVal p%) As Boolean
If p2% > 0 Then
 If p1% = p% Or p2% = p% Then
    is_point_in_element = True
 End If
ElseIf p2% = -10 Then
 If Dtwo_point_line(p1%).data(0).v_poi(0) = p% Or _
       Dtwo_point_line(p1%).data(0).v_poi(1) = p% Then
        is_point_in_element = True
 End If
ElseIf p2% = -1 Or p2% = -2 Or p2% = -3 Or p2% = -4 Or p2 = -6 Then
 is_point_in_element = is_point_in_angle(p1%, p%)
ElseIf p2% = -5 Then
 is_point_in_element = is_point_in_item(p1%, p%)
End If
End Function
Public Function is_point_in_angle(ByVal A%, ByVal p%) As Boolean
 If angle(A%).data(0).poi(1) = p% Then
    is_point_in_angle = True
 ElseIf is_point_in_line3(p%, m_lin(angle(A%).data(0).line_no(0)).data(0).data0, 0) Then
    is_point_in_angle = True
 ElseIf is_point_in_line3(p%, m_lin(angle(A%).data(0).line_no(1)).data(0).data0, 0) Then
    is_point_in_angle = True
 End If
End Function



Public Function min_positive_number(ByVal n1%, ByVal n2%) As Integer
If n1% > 0 And n2% > 0 Then
   n2% = min(n1%, n2%)
ElseIf n2% = 0 Then
   n2% = n1%
End If
min_positive_number = n2%
End Function

Public Sub set_element_depend(ByVal ty As Byte, ByVal no%, d_e1_ty As Byte, d_e1_no%, _
                 d_e2_ty As Byte, d_e2_no%, d_e3_ty As Byte, d_e3_no%, add_condition As Boolean)
Dim d, bra  As Integer
Dim d_poi(8) As Integer
If ty = point_ Then '点
    If d_e1_no% = 0 Then
       Exit Sub
    End If
     If d_e1_no% > 0 Then
        If is_depend_point_of_element(d_e1_ty, d_e1_no%, no%) Then
        d_e1_no% = 0
        d_e1_ty = 0
        End If
     End If
     If d_e2_no% > 0 Then
        If is_depend_point_of_element(d_e2_ty, d_e2_no%, no%) Then
        d_e2_no% = 0
        d_e2_ty = 0
        End If
     End If
     If d_e3_no% > 0 Then
        If is_depend_point_of_element(d_e3_ty, d_e3_no%, no%) Then
        d_e3_no% = 0
        d_e3_ty = 0
        End If
     End If
     If m_poi(no%).data(0).parent.co_degree = 2 And d_e1_no% = 0 _
            And d_e2_no% = 0 Then '初始点
        'm_poi(no%).data(0).degree = 0
        m_poi(no%).data(0).parent.co_degree = 2
        m_poi(no%).data(0).degree_for_reduce = 0
     Else
       If d_e2_no% > 0 And d_e1_no% > 0 And d_e3_no% > 0 Then
       m_poi(no%).data(0).parent.element(1).ty = d_e1_ty
       m_poi(no%).data(0).parent.element(2).ty = d_e2_ty
       m_poi(no%).data(0).parent.element(3).ty = d_e3_ty
       m_poi(no%).data(0).parent.element(1).no = d_e1_no%
       m_poi(no%).data(0).parent.element(2).no = d_e2_no%
       m_poi(no%).data(0).parent.element(3).no = d_e3_no%
       m_poi(no%).data(0).parent.co_degree = 0
       Call read_depend_point_from_element(d_e1_ty, d_e1_no%, d_poi())
       Call set_depend_point_for_element(d_poi(), point_, no%)
       Call read_depend_point_from_element(d_e2_ty, d_e2_no%, d_poi())
       Call set_depend_point_for_element(d_poi(), point_, no%)
       Call read_depend_point_from_element(d_e3_ty, d_e3_no%, d_poi())
       Call set_depend_point_for_element(d_poi(), point_, no%)
       ' m_poi(no%).data(0).degree = 0
       m_poi(no%).data(0).parent.co_degree = 2
       ElseIf d_e2_no% > 0 And d_e1_no% > 0 Then '两个依赖数据
       m_poi(no%).data(0).parent.element(1).ty = d_e1_ty
       m_poi(no%).data(0).parent.element(2).ty = d_e2_ty
       m_poi(no%).data(0).parent.element(1).no = d_e1_no%
       m_poi(no%).data(0).parent.element(2).no = d_e2_no%
       Call read_depend_point_from_element(d_e1_ty, d_e1_no%, d_poi())
       Call set_depend_point_for_element(d_poi(), point_, no%)
       Call read_depend_point_from_element(d_e2_ty, d_e2_no%, d_poi())
       Call set_depend_point_for_element(d_poi(), point_, no%)
        m_poi(no%).data(0).parent.co_degree = 0
       If d_e1_ty = point_ And d_e2_ty = point_ Then '两点
          If add_condition Then '有附加条件,比,中点等
            m_poi(no%).data(0).parent.co_degree = 2
            'm_poi(no%).data(0).degree = 0
          Else
            m_poi(no%).data(0).parent.co_degree = 1
           'm_poi(no%).data(0).degree = 1
          End If
       ElseIf (d_e1_ty = line_ Or d_e1_ty = circle_) And _
               (d_e2_ty = line_ Or d_e2_ty = circle_) Then '交点
           m_poi(no%).data(0).parent.co_degree = 2
           'm_poi(no%).data(0).degree = 0
       Else
          If add_condition Then '有附加条件,比,中点等
           m_poi(no%).data(0).parent.co_degree = 2
           'm_poi(no%).data(0).degree = 0
          Else
           m_poi(no%).data(0).parent.co_degree = 1
           'm_poi(no%).data(0).degree = 1
          End If
       End If
       ElseIf d_e2_no% > 0 Then
       m_poi(no%).data(0).parent.element(1).ty = d_e2_ty
       m_poi(no%).data(0).parent.element(1).no = d_e2_no%
       Call read_depend_point_from_element(d_e2_ty, d_e2_no%, d_poi())
       Call set_depend_point_for_element(d_poi(), point_, no%)
        m_poi(no%).data(0).parent.co_degree = 1
       If d_e2_ty = line_ Or d_e2_ty = circle_ Then '有附加条件,比,中点等
        If add_condition Then
           m_poi(no%).data(0).parent.co_degree = 2
           'm_poi(no%).data(0).degree = 0
        Else
           m_poi(no%).data(0).parent.co_degree = 1
           'm_poi(no%).data(0).degree = 1
        End If
        End If
       ElseIf d_e1_no% > 0 Then
       m_poi(no%).data(0).parent.element(1).ty = d_e1_ty
       m_poi(no%).data(0).parent.element(1).no = d_e1_no%
       Call read_depend_point_from_element(d_e1_ty, d_e1_no%, d_poi())
       Call set_depend_point_for_element(d_poi(), point_, no%)
        m_poi(no%).data(0).parent.co_degree = 1
        If d_e1_ty = line_ Or d_e1_ty = circle_ Then
         If add_condition Then
           m_poi(no%).data(0).parent.co_degree = 2
           'm_poi(no%).data(0).degree = 0
         Else
           m_poi(no%).data(0).parent.co_degree = 1
           'm_poi(no%).data(0).degree = 1
         End If
        End If
      End If
     End If
ElseIf ty = line_ Then
       If d_e1_no% > 0 And d_e2_no% > 0 Then
       d = element_depend_degree(d_e1_ty, d_e1_no%)
       d = max(element_depend_degree(d_e2_ty, d_e2_no%), d)
       'm_lin(no%).data(0).depend_poi(0) = 0
        If d_e2_no% > 0 And d_e1_no% > 0 Then
         m_lin(no%).data(0).parent.element(1).ty = d_e1_ty
         m_lin(no%).data(0).parent.element(2).ty = d_e2_ty
         m_lin(no%).data(0).parent.element(1).no = d_e1_no%
         m_lin(no%).data(0).parent.element(2).no = d_e2_no%
       Call read_depend_point_from_element(d_e1_ty, d_e1_no%, d_poi())
       Call set_depend_point_for_element(d_poi(), line_, no%)
       Call read_depend_point_from_element(d_e2_ty, d_e2_no%, d_poi())
       Call set_depend_point_for_element(d_poi(), line_, no%)
        ElseIf d_e2_no% > 0 Then
         m_lin(no%).data(0).parent.element(1).ty = d_e2_ty
         m_lin(no%).data(0).parent.element(1).no = d_e2_no%
       Call read_depend_point_from_element(d_e2_ty, d_e2_no%, d_poi())
       Call set_depend_point_for_element(d_poi(), line_, no%)
        ElseIf d_e1_no% > 0 Then
         m_lin(no%).data(0).parent.element(1).ty = d_e1_ty
         m_lin(no%).data(0).parent.element(1).no = d_e1_no%
       Call read_depend_point_from_element(d_e1_ty, d_e1_no%, d_poi())
       Call set_depend_point_for_element(d_poi(), line_, no%)
        End If
       ElseIf run_statue = 1 And m_lin(no%).data(0).data0.in_point(0) = 2 Then '12.10
        d = element_depend_degree(point_, m_lin(no%).data(0).data0.poi(0))
        d = max(element_depend_degree(point_, m_lin(no%).data(0).data0.poi(1)), d)
       m_lin(no%).data(0).parent.co_degree = d - 1
         m_lin(no%).data(0).parent.element(1).ty = point_
         m_lin(no%).data(0).parent.element(2).ty = point_
         m_lin(no%).data(0).parent.element(1).no = m_lin(no%).data(0).data0.poi(0)
         m_lin(no%).data(0).parent.element(2).no = m_lin(no%).data(0).data0.poi(1)
       End If
       If m_lin(no%).data(0).parent.element(2).no = 0 Then
        If m_lin(no%).data(0).parent.element(1).ty = point_ And _
           m_lin(no%).data(0).parent.element(2).ty = point_ Then
          If m_poi(m_lin(no%).data(0).parent.element(0).no).data(0).parent.co_degree < 2 And _
              m_poi(m_lin(no%).data(0).parent.element(1).no).data(0).parent.co_degree < 2 Then
               m_lin(no%).data(0).parent.co_degree = 2
          ElseIf m_poi(m_lin(no%).data(0).parent.element(0).no).data(0).parent.co_degree = 2 And _
              m_poi(m_lin(no%).data(0).parent.element(1).no).data(0).parent.co_degree = 2 Then
               m_lin(no%).data(0).parent.co_degree = 0
          Else
               m_lin(no%).data(0).parent.co_degree = 1
          End If
         ElseIf m_lin(no%).data(0).parent.element(1).ty = point_ And _
           m_lin(no%).data(0).parent.element(2).ty = line_ Then
            m_lin(no%).data(0).parent.co_degree = 1
         End If
       Else
          If m_lin(no%).data(0).parent.element(1).ty = point_ And _
               m_lin(no%).data(0).parent.element(2).ty = line_ And _
                (m_lin(no%).data(0).parent.element(3).ty = line_ Or _
                  m_lin(no%).data(0).parent.element(3).ty = 0) Then
             If m_poi(m_lin(no%).data(0).parent.element(0).no).data(0).parent.co_degree < 2 Then
                m_lin(no%).data(0).parent.co_degree = 1
             Else
                m_lin(no%).data(0).parent.co_degree = 0
             End If
          End If
       End If
 ElseIf ty = circle_ Then
    If m_Circ(no%).data(0).depend_element.depend_degree > 0 Then
       Exit Sub
    Else
       If d_e1_no% = 0 Then
          If m_Circ(no%).data(0).data0.center < m_Circ(no%).data(0).data0.in_point(3) Or _
             m_Circ(no%).data(0).data0.in_point(3) = 0 Then
             d_e1_ty = point_
             d_e2_ty = point_
             d_e1_no% = m_Circ(no%).data(0).data0.center
             d_e2_no% = m_Circ(no%).data(0).data0.in_point(1)
          Else
             d_e1_ty = point_
             d_e2_ty = point_
             d_e3_ty = point_
             d_e1_no% = m_Circ(no%).data(0).data0.in_point(1)
             d_e2_no% = m_Circ(no%).data(0).data0.in_point(2)
             d_e3_no% = m_Circ(no%).data(0).data0.in_point(3)
          End If
       End If
       d = element_depend_degree(d_e1_ty, d_e1_no%)
       d = max(element_depend_degree(d_e2_ty, d_e2_no%), d)
       d = max(element_depend_degree(d_e3_ty, d_e3_no%), d)
       m_Circ(no%).data(0).depend_element.depend_degree = d + 1
       m_Circ(no%).data(0).parent.element(1).ty = d_e1_ty
       m_Circ(no%).data(0).parent.element(2).ty = d_e2_ty
       m_Circ(no%).data(0).parent.element(3).ty = d_e3_ty
       m_Circ(no%).data(0).parent.element(1).no = d_e1_no%
       m_Circ(no%).data(0).parent.element(2).no = d_e2_no%
       m_Circ(no%).data(0).parent.element(3).no = d_e3_no%
       Call read_depend_point_from_element(d_e1_ty, d_e1_no%, d_poi())
       Call set_depend_point_for_element(d_poi(), circle_, no%)
       Call read_depend_point_from_element(d_e2_ty, d_e2_no%, d_poi())
       Call set_depend_point_for_element(d_poi(), circle_, no%)
       Call read_depend_point_from_element(d_e3_ty, d_e3_no%, d_poi())
       Call set_depend_point_for_element(d_poi(), circle_, no%)
End If
 End If
 End Sub
Public Function element_depend_degree(ByVal ty As Byte, ByVal no%) As Integer
If ty = point_ Then
   If m_poi(no%).data(0).depend_element.depend_degree = 0 Then
      m_poi(no%).data(0).depend_element.depend_degree = 1
   End If
   element_depend_degree = m_poi(no%).data(0).depend_element.depend_degree
ElseIf ty = line_ Then
   element_depend_degree = m_lin(no%).data(0).depend_element.depend_degree
ElseIf ty = circle_ Then
   element_depend_degree = m_Circ(no%).data(0).depend_element.depend_degree
End If
End Function

Public Function is_point_depend_conclusion_point(ByVal conc_n%, ByVal p%) As Boolean
If m_poi(p%).data(0).parent.element(0).no > 0 Then
If m_poi(p%).data(0).parent.element(1).no = 0 Then
   If m_poi(p%).data(0).parent.element(0).ty = line_ Then
      If m_lin(m_poi(p%).data(0).parent.element(0).no).data(0).parent.element(0).ty = point_ And _
           m_lin(m_poi(p%).data(0).parent.element(0).no).data(0).parent.element(1).ty = point_ Then
          If is_point_depend_conclusion_point0(conc_n%, _
              m_lin(m_poi(p%).data(0).parent.element(0).no).data(0).parent.element(0).no) And _
             is_point_depend_conclusion_point0(conc_n%, _
              m_lin(m_poi(p%).data(0).parent.element(0).no).data(0).parent.element(1).no) Then
               Call modify_depend_point(p%, line_, m_poi(p%).data(0).parent.element(0).no)
          End If
      End If
   ElseIf m_poi(p%).data(0).parent.element(0).ty = circle_ Then
   End If
End If
End If
End Function
Public Function is_point_depend_conclusion_point0(ByVal conc_n%, ByVal p%) As Boolean
Dim i%
For i% = 1 To conclusion_point(conc_n%).poi(0)
    If conclusion_point(conc_n%).poi(i%) = p% Then
       is_point_depend_conclusion_point0 = True
        Exit Function
    End If
Next
End Function
Public Sub modify_depend_point(ByVal p%, ByVal ty As Byte, ByVal no%)
Dim i% 'no%是p%依赖
For i% = 1 To m_poi(p%).data(0).in_line(0)
   If m_poi(p%).data(0).in_line(i%) <> no% And ty = line_ Then
      If m_lin(m_poi(p%).data(0).in_line(i%)).data(0).parent.element(0).ty = point_ And _
          m_lin(m_poi(p%).data(0).in_line(i%)).data(0).parent.element(0).no = p% Then
          Call modify_depend_point0(p%, line_, m_poi(p%).data(0).in_line(i%), 0)
          m_poi(p%).data(0).parent.element(1).ty = line_
          m_poi(p%).data(0).parent.element(1).no = m_poi(p%).data(0).in_line(i%)
          Exit Sub
      ElseIf m_lin(m_poi(p%).data(0).in_line(i%)).data(0).parent.element(1).ty = point_ And _
          m_lin(m_poi(p%).data(0).in_line(i%)).data(0).parent.element(1).no = p% Then
          Call modify_depend_point0(p%, line_, m_poi(p%).data(0).in_line(i%), 1)
          m_poi(p%).data(0).parent.element(1).ty = line_
          m_poi(p%).data(0).parent.element(1).no = m_poi(p%).data(0).in_line(i%)
          Exit Sub
      End If
   End If
Next
End Sub
Public Sub modify_depend_point0(ByVal p%, ByVal ty As Byte, ByVal no%, dep_no%)
Dim i%, tp% 'p%是no%依赖点
If ty = line_ Then
   For i% = 1 To m_lin(no%).data(0).data0.in_point(0)
     If m_lin(no%).data(0).data0.in_point(i%) <> m_lin(no%).data(0).parent.element(dep_no%).no Then
         If m_lin(no%).data(0).data0.in_point(i%) <> m_lin(no%).data(0).parent.element((dep_no% + 1) Mod 2).no And _
               m_lin(no%).data(0).parent.element((dep_no% + 1) Mod 2).ty = point_ Then
            If m_poi(m_lin(no%).data(0).data0.in_point(i%)).data(0).parent.element(1).no > 0 Then
                tp% = max(tp%, m_lin(no%).data(0).data0.in_point(i%))
            End If
         End If
     End If
   Next i%
If m_poi(tp%).data(0).parent.element(0).no <> no% And _
      m_poi(tp%).data(0).parent.element(1).ty = line_ Then
      m_poi(tp%).data(0).parent.element(1).no = 0
      m_poi(tp%).data(0).parent.element(1).ty = 0
      m_lin(no%).data(0).parent.element(dep_no%).ty = point_
      m_lin(no%).data(0).parent.element(dep_no%).no = tp%
ElseIf m_poi(tp%).data(0).parent.element(1).no <> no% And _
      m_poi(tp%).data(0).parent.element(0).ty = line_ Then
      m_poi(tp%).data(0).parent.element(0).no = m_poi(tp%).data(0).parent.element(1).no
      m_poi(tp%).data(0).parent.element(0).ty = m_poi(tp%).data(0).parent.element(1).ty
      m_lin(no%).data(0).parent.element(dep_no%).ty = point_
      m_lin(no%).data(0).parent.element(dep_no%).no = tp%
End If
End If
End Sub
Public Sub read_depend_point_from_element(ByVal element_type As Byte, ByVal element_no%, d_poi() As Integer)
Dim i%
If element_type = point_ Then
    If m_poi(element_no%).data(0).depend_poi(0) = 0 Then
       d_poi(0) = 1
       d_poi(1) = element_no%
    Else
     For i% = 0 To 8
      d_poi(i%) = m_poi(element_no%).data(0).depend_poi(i%)
     Next i%
    End If
ElseIf element_type = line_ Then
   If m_lin(element_no%).data(0).parent.element(0).no = 0 Then
      Call set_element_depend_for_line(element_no%)
   End If
   For i% = 0 To 8
   d_poi(i%) = m_lin(element_no%).data(0).parent.element(1).no
   Next i%
ElseIf element_type = circle_ Then
   If m_Circ(element_no%).data(0).parent.element(0).no = 0 Then
      Call set_element_depend_for_circle(element_no%)
   End If
   For i% = 0 To 8
   d_poi(i%) = m_Circ(element_no%).data(0).depend_poi(i%)
   Next i%
End If
End Sub

Public Sub set_depend_point_for_element(d_poi() As Integer, ByVal element_type As Byte, ByVal element_no%)
Dim i%, j%
If element_type = point_ Then
   For i% = 1 To d_poi(0)
    j% = 1
    Do While j% <= m_poi(element_no%).data(0).depend_poi(0)
       If m_poi(element_no%).data(0).depend_poi(j%) = d_poi(i%) Then
                   GoTo set_depend_point_for_element_mark1
       End If
       j% = j% + 1
    Loop
    m_poi(element_no%).data(0).depend_poi(0) = _
           m_poi(element_no%).data(0).depend_poi(0) + 1
    m_poi(element_no%).data(0).depend_poi(m_poi(element_no%).data(0).depend_poi(0)) = d_poi(i%)
set_depend_point_for_element_mark1:
   Next i%
       m_poi(element_no%).data(0).degree_for_reduce = m_poi(element_no%).data(0).depend_poi(0)
ElseIf element_type = line_ Then
   For i% = 1 To d_poi(0)
    j% = 1
    Do While j% <= m_lin(element_no%).data(0).parent.element(1).no
       If m_lin(element_no%).data(0).depend_poi(j%) = d_poi(i%) Then
                  GoTo set_depend_point_for_element_mark2
       End If
        j% = j% + 1
    Loop
    m_lin(element_no%).data(0).depend_poi(0) = _
           m_lin(element_no%).data(0).depend_poi(0) + 1
    m_lin(element_no%).data(0).depend_poi(m_lin(element_no%).data(0).depend_poi(0)) = d_poi(i%)
set_depend_point_for_element_mark2:
   Next i%
    m_lin(element_no%).data(0).degree = m_lin(element_no%).data(0).depend_poi(0)
ElseIf element_type = circle_ Then
   For i% = 1 To d_poi(0)
    j% = 1
    Do While j% <= m_Circ(element_no%).data(0).depend_poi(0)
       If m_Circ(element_no%).data(0).depend_poi(j%) = d_poi(i%) Then
         GoTo set_depend_point_for_element_mark3
       End If
       j% = j% + 1
    Loop
    m_Circ(element_no%).data(0).depend_poi(0) = _
           m_Circ(element_no%).data(0).depend_poi(0) + 1
    m_Circ(element_no%).data(0).depend_poi(m_Circ(element_no%).data(0).depend_poi(0)) = d_poi(i%)
set_depend_point_for_element_mark3:
   Next i%
    m_Circ(element_no%).data(0).degree = m_Circ(element_no%).data(0).depend_poi(0)
End If
End Sub
Public Sub set_element_depend_for_line(ByVal line_no%)
Call read_initail_points_for_line(m_lin(line_no%).data(0))
Call set_element_depend(line_, line_no%, _
         point_, m_lin(line_no%).data(0).data0.poi(0), _
          point_, m_lin(line_no%).data(0).data0.poi(1), _
            0, 0, False)
End Sub
Public Sub set_element_depend_for_circle(ByVal circle_no%)
If m_Circ(circle_no%).data(0).data0.center < _
     m_Circ(circle_no%).data(0).data0.in_point(3) Or _
       m_Circ(circle_no%).data(0).data0.in_point(3) = 0 Then
Call set_element_depend(circle_, circle_no%, point_, m_Circ(circle_no%).data(0).data0.center, _
             point_, m_Circ(circle_no%).data(0).data0.in_point(1), _
              0, 0, False)
Else
Call set_element_depend(circle_, circle_no%, point_, m_Circ(circle_no%).data(0).data0.in_point(1), _
             point_, m_Circ(circle_no%).data(0).data0.in_point(2), _
              point_, m_Circ(circle_no%).data(0).data0.in_point(3), False)
End If
End Sub
Public Function is_depend_point_of_element(ByVal ty As Byte, ByVal no%, ByVal p%) As Boolean
Dim d_poi(8) As Integer
Dim i%
If ty = point_ Then
 If no% = p% Then
  is_depend_point_of_element = True
 End If
Else
 Call read_depend_point_from_element(ty, no%, d_poi())
 For i% = 1 To d_poi(0)
  If d_poi(i%) = p% Then
   is_depend_point_of_element = True
    Exit Function
  End If
 Next i%
End If
End Function

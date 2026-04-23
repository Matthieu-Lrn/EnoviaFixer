6)
    dResultMatrix(10) = dMatrix1(9) * dMatrix2(1) + dMatrix1(10) * dMatrix2(4) + dMatrix1(11) * dMatrix2(7)
    dResultMatrix(11) = dMatrix1(9) * dMatrix2(2) + dMatrix1(10) * dMatrix2(5) + dMatrix1(11) * dMatrix2(8)
    
    If bAddInitialValue Then
        dResultMatrix(9) = dResultMatrix(9) + dMatrix2(9)
        dResultMatrix(10) = dResultMatrix(10) + dMatrix2(10)
        dResultMatrix(11) = dResultMatrix(11) + dMatrix2(11)
    End If
    
End Sub

Public Sub SubstractVector(ByRef dMatrix1 As Variant, ByRef dMatrix2 As Variant)
'Remove 1x3 vector2 from 1x3 vector1

    dMatrix1(9) = dMatrix1(9) - dMatrix2(9)
    dMatrix1(10) = dMatrix1(10) - dMatrix2(10)
    dMatrix1(11) = dMatrix1(11) - dMatrix2(11)
    
End Sub

Public Function GetInstanceActivity(ByVal oInstance As Product, Optional ByVal bSetDefaultMode As Boolean = False) As Boolean
'********************************************************************************
'* Name: bInstanceActivity
'* Purpose: If Component activity state is found the it returns the value else it returns false.
'*          Hence false could either mean no parameter found or instance is not active.
'*          if bSetDefaultMode is set to true then it tries to set instance in default mode if no parameter is found.
'*
'* Assumption:
'*
'* Author: Abhishek Kamboj
'* Updated by:
'* Language: VBA
'********************************************************************************
    Dim oProdparams As Parameters
    
    On Error Resume Next
    Set oProdparams = oInstance.Parameters.SubList(oInstance, False)
   
    If oProdparams.Count = 0 And bSetDefaultMode = True Then
        oInstance.ApplyWorkMode DEFAULT_MODE
        Set oProdparams = oInstance.Parameters.SubList(oInstance, False)
    End If
    
    If oProdparams.Count > 0 Then GetInstanceActivity = oProdparams.GetItem("Component Activation State").Value
    
    On Error GoTo 0
    
End Function

'****--------------------------------------------------------------------------------------------------------------------------------
'****
'****
'****                           Following are Linear algebra & Geometry algorithms functions
'****
'**** Author: Abhishek Kamboj
'**** Updated by:
'**** Language: VBA
'****--------------------------------------------------------------------------------------------------------------------------------
Public Function DotProd(v1 As Variant, v2 As Variant) As Double
'*** dot product of two vectors of 3 elements
    DotProd = v1(0) * v2(0) + v1(1) * v2(1) + v1(2) * v2(2)
End Function
Public Sub CrossProduct(v1() As Double, v2() As Double, ByRef dCrossProduct As Variant)
    Dim V(2) As Double
    
    V(0) = v1(1) * v2(2) - v1(2) * v2(1)
    V(1) = v1(2) * v2(0) - v1(0) * v2(2)
    V(2) = v1(0) * v2(1) - v1(1) * v2(0)

    dCrossProduct = V
End Sub
Public Function LengthVector(v1 As Variant) As Double
'*** Magnitude of vector of 3 elements
    LengthVector = Sqr(v1(0) ^ 2 + v1(1) ^ 2 + v1(2) ^ 2)
End Function
Public Function ArcCos(oValue As Double) As Double
'*** Inverse cosine of value
    If Round(oValue, 8) = 1 Then ArcCos = 0: Exit Function
    If Round(oValue, 8) = -1 Then ArcCos = PI: Exit Function
    ArcCos = Atn(-oValue / Sqr(1 - oValue ^ 2)) + 2 * Atn(1)
End Function

Public Function AngleInRad(ByVal dMatrix1 As Variant, ByVal dMatrix2 As Variant, Optional ByRef sVector As String = "X") As Double
'*** Find angle between two vectors in 3 dimensions.Input is position matrix of two instance
'*** By assigning optional value to SVector , one can find angle between either X ,Y or Z axis of two Position matrix.
    Dim Vctr1(2)
    Dim Vctr2(2)
    
    Vctr1(0) = dMatrix1(0)
    Vctr1(1) = dMatrix1(1)
    Vctr1(2) = dMatrix1(2)
    
    Vctr2(0) = dMatrix2(0)
    Vctr2(1) = dMatrix2(1)
    Vctr2(2) = dMatrix2(2)
    
    If sVector = "Y" Then
        Vctr1(0) = dMatrix1(3)
        Vctr1(1) = dMatrix1(4)
        Vctr1(2) = dMatrix1(5)
        
        Vctr2(0) = dMatrix2(3)
        Vctr2(1) = dMatrix2(4)
        Vctr2(2) = dMatrix2(5)
    ElseIf sVector = "Z" Then
        Vctr1(0) = dMatrix1(6)
        Vctr1(1) = dMatrix1(7)
        Vctr1(2) = dMatrix1(8)
        
        Vctr2(0) = dMatrix2(6)
        Vctr2(1) = dMatrix2(7)
        Vctr2(2) = dMatrix2(8)
    End If
    
    AngleInRad = ArcCos(DotProd(Vctr1, Vctr2) / (LengthVector(Vctr1) * LengthVector(Vctr2)))
    
End Function
Public Function dDistanceLineSegToLineSeg(p1() As Double, p2() As Double, p3() As Double, p4() As Double) As Double

'********--------------------
'***
'*** returns shortest distance between two line segments Line 1 is from P1 to P2 and Line 2 is from P3 to P4
'*** Line 1 is from P1 to P2 and Line 2 is from P3 to P4
'*** Adapted from http://geomalgorithms.com/a07-_distance.html#dist3D_Segment_to_Segment()
'***
'********--------------------
    Dim U(2) As Double
    Dim V(2) As Double
    Dim W(2) As Double
    Dim Result(2) As Double
    Dim PR1(2) As Double
    Dim PR2(2) As Double
    Dim a As Double, b As Double, c As Double, d As Double, DD As Double, sc As Double, sn As Double, sd As Double
    Dim tc As Double, tN As Double, tD As Double
    
    U(0) = p2(0) - p1(0): U(1) = p2(1) - p1(1): U(2) = p2(2) - p1(2)
    V(0) = p4(0) - p3(0): V(1) = p4(1) - p3(1): V(2) = p4(2) - p3(2)
    W(0) = p1(0) - p3(0): W(1) = p1(1) - p3(1): W(2) = p1(2) - p3(2)

    a = DotProd(U, U)   ' should be > = zero
    b = DotProd(U, V)
    c = DotProd(V, V)   ' should be > = zero
    d = DotProd(U, W)
    E = DotProd(V, W)
    
    DD = a * c - b * b
    sd = DD ' default
    tD = DD ' default
    
    If DD < SMALLNUMBER Then     ' lines are parallel
            sn = 0               ' forcing use of point P1 of Line 1
            sd = 1               ' to prevent possible division by 0.0 later
            tN = E
            tD = c
    Else                         ' if lines are not parallel then get closest point on infinite lines
        sn = (b * E - c * d)
        tN = (a * E - b * d)
            If sn < 0 Then       ' sc < 0 => the s=0 edge is visible
                sn = 0
                tN = E
                tD = c
            ElseIf sn > sd Then  ' sc > 1  => the s=1 edge is visible
                sn = sd
                tN = E + b
                tD = c
            End If
    End If
    If (tN < 0) Then             'tc < 0 => the t=0 edge is visible
        tN = 0
        ' recompute sc for this edge
            If (0 - d) < 0 Then
                sn = 0
            ElseIf (0 - d) > a Then
                sn = sd
            Else
                sn = 0 - d
                sd = a
            End If
    ElseIf tN > tD Then         ' tc > 1  => the t=1 edge is visible
        tN = tD
        'recompute sc for this edge
            If (0 - d + b) < 0 Then
                sn = 0
            ElseIf (0 - d + b) > a Then
                sn = sd
            Else
                sn = (0 - d + b)
                sd = a
            End If
    End If
        
    If Abs(sn) < SMALLNUMBER Then
        sc = 0
    Else
        sc = sn / sd
    End If
    
    If Abs(tN) < SMALLNUMBER Then
        tc = 0
    Else
        tc = tN / tD
    End If
    Result(0) = W(0) + (sc * U(0)) - (tc * V(0))
    Result(1) = W(1) + (sc * U(1)) - (tc * V(1))
    Result(2) = W(2) + (sc * U(2)) - (tc * V(2))
    dDistanceLineSegToLineSeg = LengthVector(Result) ' Length of Vector
    PR1(0) = p1(0) + sc * U(0)
    PR1(1) = p1(1) + sc * U(1)
    PR1(2) = p1(2) + sc * U(2)
    
    PR2(0) = p3(0) + tc * V(0)
    PR2(1) = p3(1) + tc * V(1)
    PR2(2) = p3(2) + tc * V(2)
End Function
Public Function dDistancePointToLineSeg(p1() As Double, p2() As Double, p3() As Double, Optional LineIsInfinite As Boolean = False) As Double
'*** calculate shortest distance between a point & a line segment
    Dim U(2) As Double
    Dim V(2) As Double
    Dim W(2) As Double
    Dim a As Double, b As Double, c As Double
    Dim ClosestPoint(2) As Double
    Dim Result(2) As Double
    
    U(0) = p1(0) - p3(0): U(1) = p1(1) - p3(1): U(2) = p1(2) - p3(2)     ' Vector of End point of line to the Point
    V(0) = p3(0) - p2(0): V(1) = p3(1) - p2(1): V(2) = p3(2) - p2(2)     ' Vector of Line
    W(0) = p1(0) - p2(0): W(1) = p1(1) - p2(1): W(2) = p1(2) - p2(2)     ' Vector of Startpoint of Line to the point
        
    a = DotProd(V, W)
    b = DotProd(V, V)
    
    If Not LineIsInfinite Then  ' conditions for line of finite length
        If a < 0 Then
            dDistancePointToLineSeg = LengthVector(W)
            Exit Function
        End If
        If b < a Then
            dDistancePointToLineSeg = LengthVector(U)
            Exit Function
        End If
    End If
    ' if line is infinte
    c = a / b
    ClosestPoint(0) = p2(0) + c * V(0)
    ClosestPoint(1) = p2(1) + c * V(1)
    ClosestPoint(2) = p2(2) + c * V(2)
    
    Result(0) = p1(0) - ClosestPoint(0)
    Result(1) = p1(1) - ClosestPoint(1)
    Result(2) = p1(2) - ClosestPoint(2)
    
    dDistancePointToLineSeg = Round(LengthVector(Result), 12)
End Function
Public Function dDistanceLineToLine(p1() As Double, p2() As Double, p3() As Double, p4() As Double) As Double
'*** calculates shortest distance between lines of infinite length
    Dim U(2) As Double
    Dim V(2) As Double
    Dim W(2) As Double
    Dim PR1(2) As Double
    Dim PR2(2) As Double
    
    Dim Result(2) As Double
    Dim a As Double, b As Double, c As Double, d As Double, DD As Double, sc As Double, tc As Double
    
    U(0) = p2(0) - p1(0): U(1) = p2(1) - p1(1): U(2) = p2(2) - p1(2)    ' vector of line 1
    V(0) = p4(0) - p3(0): V(1) = p4(1) - p3(1): V(2) = p4(2) - p3(2)    ' vector of line 2
    W(0) = p1(0) - p3(0): W(1) = p1(1) - p3(1): W(2) = p1(2) - p3(2)

    a = DotProd(U, U)   ' should be > = zero
    b = DotProd(U, V)
    c = DotProd(V, V)   ' should be > = zero
    d = DotProd(U, W)
    E = DotProd(V, W)
    
    DD = a * c - b * b
     If DD < SMALLNUMBER Then     ' lines are parallel
        sc = 0
        If b > c Then
            tc = d / b
        Else
            tc = E / c
        End If
    Else
        sc = ((b * E) - (c * d)) / DD
        tc = ((a * E) - (b * d)) / DD
    End If

    Result(0) = W(0) + (sc * U(0)) - (tc * V(0))
    Result(1) = W(1) + (sc * U(1)) - (tc * V(1))
    Result(2) = W(2) + (sc * U(2)) - (tc * V(2))

    dDistanceLineToLine = Round(LengthVector(Result), 12)
    
    PR1(0) = p1(0) + sc * U(0)
    PR1(1) = p1(1) + sc * U(1)
    PR1(2) = p1(2) + sc * U(2)
    
    PR2(0) = p3(0) + tc * V(0)
    PR2(1) = p3(1) + tc * V(1)
    PR2(2) = p3(2) + tc * V(2)
    
End Function
Public Function dDistancePointToPoint(p1() As Double, p2() As Double) As Double
    Dim ResultVector(2) As Double
    
    ResultVector(0) = p2(0) - p1(0): ResultVector(1) = p2(1) - p1(1): ResultVector(2) = p2(2) - p1(2)
    dDistancePointToPoint = Round(LengthVector(ResultVector), 12)
End Function
Public Function dDistancePointToPlane(p1() As Double, p2() As Double, p3() As Double) As Double
    '*** p1 is the point from which we want to measure the distance to plane
    '*** Plane vector is created from p2 & p3 , p2 is also a point lying on a plane
    '*** distance is +ve when point is on the side of vector of plane
    '*** and -ve when it is on the other side
    Dim PlaneVector(2) As Double
    Dim Result(2) As Double
    Dim sb As Double
    Dim sn As Double
    Dim sd As Double
    Dim U(2) As Double
    
    PlaneVector(0) = p3(0) - p2(0): PlaneVector(1) = p3(1) - p2(1): PlaneVector(2) = p3(2) - p2(2)
    U(0) = p1(0) - p2(0): U(1) = p1(1) - p2(1): U(2) = p1(2) - p2(2)
    
    sn = DotProd(PlaneVector, U)
    sd = LengthVector(PlaneVector)
    dDistancePointToPlane = sn / sd
End Function
'Abhishek Update with InsertHole
Public Function iFindLineSegmentIntersectionToPlane(p1() As Double, p2() As Double, p3() As Double, p4() As Double, Optional ByRef dIntsctPT As Variant) As Integer
'***-------------
'*** p1 & p2 belongs to start & end of line segment
'*** Plane vector is created from p3 & p4 , p3 being the point lying on a plane
'*** Result = 0 No intersection
'*** Result = 1 Intersection exist
'*** Result =2 segement lies on plane
'*** if Intersection exist then dIntsctPT will have coordinates of intersection
'*** Adapted from --- http://geomalgorithms.com/a05-_intersect-1.html
'***-------------
    Dim PlaneVector(2) As Double
    Dim U(2) As Double
    Dim W(2) As Double
    Dim d As Double
    Dim N As Double
    Dim SI As Double
    Dim Result As Integer
    
    PlaneVector(0) = p4(0) - p3(0): PlaneVector(1) = p4(1) - p3(1): PlaneVector(2) = p4(2) - p3(2)
    
    U(0) = p2(0) - p1(0): U(1) = p2(1) - p1(1): U(2) = p2(2) - p1(2)
    W(0) = p1(0) - p3(0): W(1) = p1(1) - p3(1): W(2) = p1(2) - p3(2)
    
    d = DotProd(PlaneVector, U)
    N = -DotProd(PlaneVector, W)
    
    If Abs(d) < SMALLNUMBER Then
        If N = 0 Then
            Result = 2 ' segment lies in plane
        Else
            Result = 0 ' no intersection
        End If
    Else
        SI = Round(N / d, 3)
        If SI < 0 Or SI > 1 Then
            Result = 0  ' no intersection
        Else
            Result = 1 ' Intersection exist
            'dIntsctPT(0) = p1(0) + SI * U(0)
            'dIntsctPT(1) = p1(1) + SI * U(1)
            'dIntsctPT(2) = p1(2) + SI * U(2)
        End If
    End If
    iFindLineSegmentIntersectionToPlane = Result
End Function
Public Function GetAttributesOfDocRevision(ByRef sAttributes As Variant, ByVal sPartNumber As String, Optional ByVal oProduct As Product = Nothing, Optional ByVal sDocRev As String = "") As String

    Dim sPreviousRevision As String
    Dim sRevision As String
    Dim oResults 'As clsCollection - clsCollection in DDP Toolbar is different from clsCollection in BDIWebServices
    Dim i As Integer
    Dim sRev As String
    Dim iLatestRev As Integer 'The revision is transformed in an integer in order to find the hightest one
    Dim sLatestRev As String: sLatestRev = ""
    Dim sFileName As String: sFileName = ""
    
    '*** If oProduct is argument is not nothing
    If Not oProduct Is Nothing Then
        On Error Resume Next
        sFileName = oProduct.ReferenceProduct.Parent.Name
        On Error GoTo 0
        'Get revision from sFileName for CATPart and SP Collector ---- Special case we don't want to use webservice in this case
        If sFileName Like "*.CATPart" Or sFileName Like "*-###SP*.CATProduct" Then
            sLatestRev = Right(Split(sFileName, ".CAT")(0), 2)
        End If
        
    ElseIf sDocRev <> "" Then
        sLatestRev = sDocRev
    End If

    If sLatestRev = "" Then
        'Search all document where the part number like sPartNumber*
        Set oResults = WebServiceAccessTool.GetDocumentByBaseNumber(sPartNumber)
    
        'No result found
        If oResults.Count = 0 Then
            Exit Function
        End If
        'Scan all results and find the document with the latest revision
        sLatestRev = "--"
        iLatestRev = 5 * Asc(Left(sLatestRev, 1)) + Asc(Right(sLatestRev, 1))
        For i = 1 To oResults.Count
            If oResults.GetItem(i).GetItem("FIELD_PART_NUMBER") = sPartNumber Then
                sRev = oResults.GetItem(i).GetItem("FIELD_DOCUMENT_REVISION")
                If 5 * Asc(Left(sRev, 1)) + Asc(Right(sRev, 1)) > iLatestRev Then
                    sLatestRev = sRev
                    iLatestRev = 5 * Asc(Left(sLatestRev, 1)) + Asc(Right(sLatestRev, 1))
                End If
            End If
        Next
    End If
    
    'Send attributes of sLatestRev
    Set sAttributes = WebServiceAccessTool.GetENOVIADocumentAttributs(sPartNumber, sLatestRev)
    GetAttributesOfDocRevision = sLatestRev

End Function




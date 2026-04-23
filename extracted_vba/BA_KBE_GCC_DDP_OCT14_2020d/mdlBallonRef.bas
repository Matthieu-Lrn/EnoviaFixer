Attribute VB_Name = "mdlBallonRef"
Option Explicit

'********************************************************************************
'* Purpose: Fill infos from a 2DLine inside a Balloon objects after selection   *
'*          of a balloon and a 2DLine. Info is FS, BL or WL direction +         *
'*          before AutoStamp operation occurs. Works for Global 7000 8000       *
'*          taking in to consideration the plug vales.Adapted from IBM's        *
'*          BABalloon ref macro.                                                *
'*                                                                              *
'* Assumption:                                                                  *
'*                                                                              *
'* Updated:  Abhishek Kamboj                                                    *
'*                                                                              *
'* Language: VBA                                                                *
'********************************************************************************

Private sDecimalPlace As String

Sub CreateFSBallon(ByVal sdummy As String)

    On Error Resume Next
    Err.Clear
    Dim oDrawingDoc As Document
    Set oDrawingDoc = CATIA.ActiveDocument
    If Err.Number <> 0 Then
        MsgBox "There is no Active Document open!" & vbCrLf & "Please open a CATDrawing document first and restart macro!", vbOKOnly, "Execution aborted !"
        bStopMultiSelection = True  '***To Exit the loop in frmBallons
        GoTo ExitProgram
    End If
    On Error GoTo 0
    
    Dim oViewToWorkIn As DrawingView
    Dim oViewToWorkInGenBehav 'As DrawingViewGenerativeBehavior
    Dim sPlane As String: sPlane = ""
    Dim oSelection 'As Selection
    Set oSelection = oDrawingDoc.Selection
    Dim vOption As VbMsgBoxResult
    Dim sFilter(0)
    Dim sStatus As String
    Dim oSelectedElement 'As SelectedElement
    Dim oDwgCompo As DrawingComponent
    Dim oLine2D 'As Line2D
    Dim Xh, Yh, Zh, Xv, Yv, Zv, Xn, Yn, Zn As Double
    Dim oStartPoint2D 'As Point2D
    Dim oEndPoint2D
    Dim dStartCoord2D(1)
    Dim dEndCoord2D(1)
    Dim dStartCoord(2)
    Dim dEndCoord(2)
    Dim dBalloonValue As Double
    Dim sBalloonValue As String
    Dim oDwgText As DrawingText
    Dim sFontName As String
    Dim dFontsize As Double
    
    Dim sXdir As String     'X Direction
    Dim sX As String        'X Sign
    Dim sYdir As String     'Y Direction
    Dim sY As String        'Y Sign
    Dim sZdir As String     'Z Direction
    Dim sZ As String        'Z Sign

    Select Case frmBalloons.lstProgram.Value
        Case "Others"
            sXdir = "FS"
            sX = "+"
            sYdir = "BL"
            sY = "+"
            sZdir = "WL"
            sZ = "+"
        Case Else
            sXdir = "BL"
            sX = "+"
            sYdir = "WL"
            sY = "+"
            sZdir = "FS"
            sZ = "+"
    End Select
    
    Set oSelectedItem = MultiWindowSelectElement("Please select a Reference balloon component")
    If oSelectedItem Is Nothing Then
        CATIA.StatusBar = ""
        GoTo ExitProgram
    End If
    If TypeName(oSelection.Item(1).Value) = "DrawingComponent" Then
        Set oSelectedElement = oSelection.Item(1)
        Set oDwgCompo = oSelectedElement.Value
        
        Do
            Set oSelectedItem = MultiWindowSelectElement("Please select the reference line used to fill reference balloon component text")
            If oSelectedItem Is Nothing Then
                CATIA.StatusBar = ""
                GoTo ExitProgram
            End If
        Loop Until TypeName(oSelection.Item(1).Value) = "Line2D"
        
        Set oSelectedElement = oSelection.Item(1)
        Set oLine2D = oSelectedElement.Value
        Set oStartPoint2D = oLine2D.StartPoint
        Set oEndPoint2D = oLine2D.EndPoint
        oStartPoint2D.GetCoordinates dStartCoord2D
        oEndPoint2D.GetCoordinates dEndCoord2D
        
        'View definition : projection plane
        Set oViewToWorkIn = oLine2D.Parent.Parent
        Debug.Print "oViewToWorkIn.Name = " & oViewToWorkIn.Name
        Set oViewToWorkInGenBehav = oViewToWorkIn.GenerativeBehavior
        oViewToWorkInGenBehav.GetProjectionPlane Xh, Yh, Zh, Xv, Yv, Zv
        '***Set Values From Form
        Dim sRef As String
        If frmBalloons.optBasic Then
            sRef = ""
        End If
        If frmBalloons.optRef Then
            sRef = vbLf & "REF"
        End If
        sDecimalPlace = frmBalloons.Cmb_Precision.Value
        'Projection plane normal vector definition
        Xn = Format(Yh * Zv - Zh * Yv, sDecimalPlace)
        Yn = Format(Xv * Zh - Xh * Zv, sDecimalPlace)
        Zn = Format(Xh * Yv - Yh * Xv, sDecimalPlace)
        'Check if plane is // (I,J,K) 3D vectors
        If Xn = 0 Then
            If Yn = 0 Then
                sPlane = "XY"
            ElseIf Zn = 0 Then
                sPlane = "XZ"
            Else
                sPlane = "X"    'Plane is parallel to (OX)
            End If
        ElseIf Yn = 0 Then
            If Zn = 0 Then
                sPlane = "YZ"
            Else
                sPlane = "Y"    'Plane is parallel to (OY)
            End If
        ElseIf Zn = 0 Then
            sPlane = "Z"    'Plane is parallel to (OZ)
        Else
            sPlane = "OTHER"
        End If
                
        Debug.Print "sPlane = " & sPlane
        
        '2D Line End Points 3D coordinates
        dStartCoord(0) = Format((dStartCoord2D(0) * Xh + dStartCoord2D(1) * Xv) / 25.4, sDecimalPlace)
        dStartCoord(1) = Format((dStartCoord2D(0) * Yh + dStartCoord2D(1) * Yv) / 25.4, sDecimalPlace)
        dStartCoord(2) = Format((dStartCoord2D(0) * Zh + dStartCoord2D(1) * Zv) / 25.4, sDecimalPlace)
        dEndCoord(0) = Format((dEndCoord2D(0) * Xh + dEndCoord2D(1) * Xv) / 25.4, sDecimalPlace)
        dEndCoord(1) = Format((dEndCoord2D(0) * Yh + dEndCoord2D(1) * Yv) / 25.4, sDecimalPlace)
        dEndCoord(2) = Format((dEndCoord2D(0) * Zh + dEndCoord2D(1) * Zv) / 25.4, sDecimalPlace)
        Debug.Print "dStartCoord(0) = " & dStartCoord(0)
        Debug.Print "dStartCoord(1) = " & dStartCoord(1)
        Debug.Print "dStartCoord(2) = " & dStartCoord(2)
        Debug.Print "dEndCoord(0) = " & dEndCoord(0)
        Debug.Print "dEndCoord(1) = " & dEndCoord(1)
        Debug.Print "dEndCoord(2) = " & dEndCoord(2)
        
        'Analyze 2D Line StartPoint and EndPoint 3D coordinates to find which info is to be copied inside Balloon
        Dim sAxis As String
        If dStartCoord(0) = dEndCoord(0) Then
            If dStartCoord(1) = dEndCoord(1) Then
                If dStartCoord(2) = dEndCoord(2) Then
                    sAxis = "Segment length equals zero"
                ElseIf sPlane = "XZ" Then
                    sAxis = "X"
                ElseIf sPlane = "YZ" Then
                    sAxis = "Y"
                Else
                    sAxis = "Segment is orthogonal to " & sXdir & " AND " & sYdir & " direction but it is not in (XZ) or (YZ) plane."   'Projection plane is perpendicular to (XY) plane, but different from (XZ) and (YZ) planes
                End If
            ElseIf dStartCoord(2) = dEndCoord(2) Then
                If sPlane = "XY" Then
                    sAxis = "X"
                ElseIf sPlane = "YZ" Then
                    sAxis = "Z"
                Else
                    sAxis = "Segment is orthogonal to " & sXdir & " AND " & sZdir & " direction but is not in (XY) or (YZ) plane."   'Projection plane is perpendicular to (XZ) plane, but different from (XY) and (YZ) planes
                End If
            ElseIf sPlane = "YZ" Then
                sAxis = "Segment is only orthogonal to " & sXdir & " direction. But this direction is orthogonal to the drawing plane."
            Else
                sAxis = "X"
            End If
        ElseIf dStartCoord(1) = dEndCoord(1) Then
            If dStartCoord(2) = dEndCoord(2) Then
                If sPlane = "XY" Then
                    sAxis = "Y"
                ElseIf sPlane = "XZ" Then
                    sAxis = "Z"
                Else
                    sAxis = "Segment is orthogonal to " & sYdir & " AND " & sZdir & " direction but is not in (XY) or (XZ) plane."   'Projection plane is perpendicular to (YZ) plane, but different from (XY) and (XZ) planes
                End If
            ElseIf sPlane = "ZX" Then
                 sAxis = "Segment is only orthogonal to " & sYdir & " direction. But this direction is orthogonal to the drawing plane."
            Else
                sAxis = "Y"
            End If
        ElseIf dStartCoord(2) = dEndCoord(2) Then
            If sPlane = "XY" Then
                 sAxis = "Segment is only orthogonal to " & sZdir & " direction. But this direction is orthogonal to the drawing plane."
            Else
                sAxis = "Z"
            End If
        Else
            sAxis = "Segment is not orthogonal to any principal direction (FS, BL or WL)."   'Selected line is not in any principal direction
        End If
        
        Select Case sAxis
            Case "X"
                If sX = "+" Then
                    dBalloonValue = dStartCoord(0)
                Else
                    dBalloonValue = -dStartCoord(0)
                End If
                Dim dPr As Long
                dPr = IIf(frmBalloons.Cmb_Precision.Value = "0.00", 2, 3)
                
                If Round(dBalloonValue, dPr) < 0 Then
                    sXdir = "LBL"
                ElseIf Round(dBalloonValue, dPr) > 0 Then
                    sXdir = "RBL"
                Else
                    sXdir = "BL"
                End If
                sBalloonValue = sXdir & vbLf & Format(Abs(dBalloonValue), sDecimalPlace) & sRef
            Case "Y"
                If sY = "+" Then
                    dBalloonValue = dStartCoord(1)
                Else
                    dBalloonValue = -dStartCoord(1)
                End If
                sBalloonValue = sYdir & vbLf & Format(dBalloonValue, sDecimalPlace) & sRef
            Case "Z"
                If sZ = "+" Then
                    dBalloonValue = dStartCoord(2)
                Else
                    dBalloonValue = -dStartCoord(2)
                End If
                sBalloonValue = sZdir & vbLf & sFuselagePlugCorrection(dBalloonValue) & sRef
            Case Else
                MsgBox sAxis & vbCrLf & "Script cannot fill the balloon !", vbOKOnly, "Balloon Information not updated!"
                GoTo ExitProgram      'To avoid text already written inside Balloon to be erased
        End Select
        
        Set oDwgText = oDwgCompo.GetModifiableObject(1)
        sFontName = oDwgText.GetFontName(1, 1)
        dFontsize = oDwgText.GetFontSize(1, 1)
        oDwgText.Text = sBalloonValue
        oDwgText.SetFontName 1, Len(sBalloonValue), sFontName
        oDwgText.SetFontSize 1, Len(sBalloonValue), dFontsize
        oDwgText.SetParameterOnSubString catAlignment, 0, 0, 1  '***Text is centered
    End If
ExitProgram:
If Err.Number <> 0 Then frmBalloons.bfrmBalloonBusy = False
End Sub
Private Function sFuselagePlugCorrection(ByVal dBallonValue As Double) As String
    Dim sReturn As String
    Dim dTemp As Double
    
    If frmBalloons.lstProgram.Value = "G7000" Then
        If dBallonValue <= 307 Then
            sReturn = Format(dBallonValue + 72, sDecimalPlace)
        ElseIf dBallonValue > 307 And dBallonValue <= 379 Then
            dTemp = dBallonValue - 307
            sReturn = Format(379, sDecimalPlace) & vbLf & "+" & Format(dTemp, sDecimalPlace)
        ElseIf dBallonValue > 813 And dBallonValue <= 849 Then
            dTemp = dBallonValue - 813
            sReturn = Format(813, sDecimalPlace) & vbLf & "+" & Format(dTemp, sDecimalPlace)
        ElseIf dBallonValue > 849 Then
            sReturn = Format(dBallonValue - 36, sDecimalPlace)
        Else
            sReturn = Format(dBallonValue, sDecimalPlace)
        End If
    Else
        sReturn = Format(dBallonValue, sDecimalPlace)
    End If
    sFuselagePlugCorrection = sReturn
End Function



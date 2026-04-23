Attribute VB_Name = "frmLocationInfo"
Attribute VB_Base = "0{32CD5167-602C-494F-A55E-BB93F0FB4CBA}{547F7BC2-9C3F-444B-9DEB-3275D2AB0301}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private oListOfCallouts As Collection
Private oListOfVueInfo As Collection
Private oListofLocationInfo As Collection
Private sWindow As Window
Private CatDoc As Document
Private oExcel As clsExcel
Private oWorksheet
Private sSourceWindow As Window

Public Sub LaunchDisplayInCatia(ByRef oExcl As clsExcel, _
                                ByRef owrksht, _
                                ByRef oCatDc As Document, _
                                ByRef oListOfCall As Collection, _
                                ByRef oListOfView As Collection, _
                                ByRef oListofLocation As Collection, _
                                ByRef sSourcewdw As Window)

    Set oListOfCallouts = oListOfCall
    Set oListOfVueInfo = oListOfView
    Set oListofLocationInfo = oListofLocation
    Set oCatDOc = oCatDc
    Set oExcel = oExcl
    Set oWorksheet = owrksht
    Set sSourceWindow = sSourcewdw
    With Me
        .Height = dTitleBarHeight + 74
        .Width = 135
        .imgDSGlobe.Top = .Height - dTitleBarHeight - 28
        .imgDSGlobe.Left = -3
        .cmdSelect.Top = .Height - dTitleBarHeight - 65
        .cmdSelect.Left = .Width - .cmdSelect.Width - 60
        .cmdCreate.Top = .cmdSelect.Top
        .cmdCreate.Left = .Width - .cmdCreate.Width - 10
        .cmdNext.Left = .cmdCreate.Left
        .cmdNext.Top = .cmdCreate.Top + .cmdCreate.Height + 5
        .chkGreen.Left = .cmdSelect.Left
        .chkGreen.Top = .cmdCreate.Top + .cmdCreate.Height + cmdNext.Height + 5
        .chkPink.Left = .cmdSelect.Left + 45
        .chkPink.Top = .chkGreen.Top
        .StartUpPosition = Manual
        .Left = dPointsToPixelRatioH * (CATIA.Left + CATIA.Width) - .Width - 85
        .Top = dPointsToPixelRatioV * (CATIA.Top + CATIA.Height) - .Height - 85
        
        .cmdSelect.SetFocus
        .Show vbModless
    End With

End Sub

Private Sub cmdCancel_Click()
    Call frmKBEMain.resetToolbar
End Sub

Private Sub chkGreen_Click()
    If Me.chkGreen.Value = True And Me.chkPink.Value = True Then
        Me.chkPink.Value = False
    End If
End Sub
Private Sub chkPink_Click()
    If Me.chkGreen.Value = True And Me.chkPink.Value = True Then
        Me.chkGreen.Value = False
    End If
End Sub

Private Sub cmdNext_Click()
    Dim iRow As Long, iCol As Long, iter As Integer, k As Integer, l As Integer
    Dim iCase As Integer
    
    Dim oSht As DrawingSheet
    Dim oColl As Collection
    Dim oText As DrawingText, vueText As DrawingText, callouttext As DrawingText
    Dim oVue As DrawingView
    Dim oRefText As DrawingText
    Dim oCurrentItem As Collection, oAssociatedView As Collection
    Dim oTextCollection As Collection
    Dim dTextwidth As Double, dTextHeight As Double
    Dim oExistingLocation As Collection
    Dim oPointCol As Collection
    Dim dSheetX As Double, dSheetY As Double
    Dim lerr As Long
    
    Dim pText(1) As Double, p1(1) As Double, p2(1) As Double, p3(1) As Double, p4(1) As Double
    Dim pLocal As Variant
    Dim sText As String
    Dim iLocation As Integer
    Dim bModified As Boolean
    
    On Error Resume Next
        iRow = oExcel.App.ActiveCell.Row
        iCol = oExcel.App.ActiveCell.Column
        lerr = Err.Number
    On Error GoTo 0
    If lerr <> 0 Then Exit Sub  'Cannot do if we cannot get excel object
    
    iCase = 0
    If oWorksheet.Cells(iRow, 1) Like "ListOfCallout;*" Then    '***Can create only if first column has value "ListOfCallout" otherwise we have insufficient info
        Set oColl = oListOfCallouts
        iCase = 1
    Else
        Exit Sub
    End If
    
    If Not IsNumeric(Split(oWorksheet.Cells(iRow, 1), ";")(UBound(Split(oWorksheet.Cells(iRow, 1), ";")))) Then Exit Sub
    If (UBound(Split(oWorksheet.Cells(iRow, 1), ";"))) > 0 Then
        iter = CInt(Split(oWorksheet.Cells(iRow, 1), ";")(1))
    Else
        Exit Sub
    End If
    If (UBound(Split(oWorksheet.Cells(iRow, 1), ";"))) = 2 Then
        iLocation = CInt(Split(oWorksheet.Cells(iRow, 1), ";")(2))
        If oWorksheet.Cells(iRow, 3) Like "DETAIL*" Then
            If iLocation >= 4 Then iLocation = 1 Else iLocation = iLocation + 1
        Else
            If iLocation = 8 Then iLocation = 1 Else iLocation = iLocation + 1
        End If
    Else
        iLocation = 1
    End If
        
    'End If
    Set oCurrentItem = oColl.Item(iter)
    Set oAssociatedView = oCurrentItem.Item("ASSOCIATEDVIEW")   '*** This should exist in List of callout
    If iCol < 7 Then
        '*** Position Existinglocation text
        If bExistInCol(oCurrentItem, "EXISTINGLOCATIONTXT") Then
            Set oExistingLocation = oCurrentItem.Item("EXISTINGLOCATIONTXT")
            Set oText = oExistingLocation.Item("TEXTOBJECT").Item(1)
            Set oTextCollection = oCurrentItem.Item("TEXTOBJECT")
            Set oPointCol = oCurrentItem.Item("POINTCOLLECTION")
            Select Case iLocation
                Case 1  'Top
                    Set oRefText = oTextCollection.Item(1)
                    p1(0) = oPointCol.Item(1).Item(1)(0): p1(1) = oPointCol.Item(1).Item(1)(1)
                    p2(0) = oPointCol.Item(1).Item(2)(0): p2(1) = oPointCol.Item(1).Item(2)(1)
                    pText(0) = p1(0) + (p2(0) - p1(0)) / 2
                    pText(1) = p1(1) + dDWGTextGAP
                    pLocal = dGetLocalPositionfromGlobal(pText, oText)
                    oText.AnchorPosition = catBottomCenter
                    oText.X = pLocal(0): oText.Y = pLocal(1)
                    Call changeTextProperties(oText)
                    bModified = True
                Case 2  'Right
                    Set oRefText = oTextCollection.Item(1)
                    p2(0) = oPointCol.Item(1).Item(2)(0): p2(1) = oPointCol.Item(1).Item(2)(1)
                    p3(0) = oPointCol.Item(1).Item(3)(0): p3(1) = oPointCol.Item(1).Item(3)(1)
                    pText(0) = p2(0) + dDWGTextGAP + 5
                    pText(1) = p3(1) + (p2(1) - p3(1)) / 2
                    pLocal = dGetLocalPositionfromGlobal(pText, oText)
                    oText.AnchorPosition = catMiddleLeft
                    oText.X = pLocal(0): oText.Y = pLocal(1)
                    Call changeTextProperties(oText)
                    bModified = True
                Case 3  'Bottom
                    Set oRefText = oTextCollection.Item(1)
                    p3(0) = oPointCol.Item(1).Item(3)(0): p3(1) = oPointCol.Item(1).Item(3)(1)
                    p4(0) = oPointCol.Item(1).Item(4)(0): p4(1) = oPointCol.Item(1).Item(4)(1)
                    pText(0) = p4(0) + (p3(0) - p4(0)) / 2
                    pText(1) = p3(1) - dDWGTextGAP - 2
                    pLocal = dGetLocalPositionfromGlobal(pText, oText)
                    oText.AnchorPosition = catTopCenter
                    oText.X = pLocal(0): oText.Y = pLocal(1)
                    Call changeTextProperties(oText)
                    bModified = True
                Case 4 'Left
                    Set oRefText = oTextCollection.Item(1)
                    p1(0) = oPointCol.Item(1).Item(1)(0): p1(1) = oPointCol.Item(1).Item(1)(1)
                    p4(0) = oPointCol.Item(1).Item(4)(0): p4(1) = oPointCol.Item(1).Item(4)(1)
                    pText(0) = p1(0) - dDWGTextGAP - 5
                    pText(1) = p4(1) + (p1(1) - p4(1)) / 2
                    pLocal = dGetLocalPositionfromGlobal(pText, oText)
                    oText.AnchorPosition = catMiddleRight
                    oText.X = pLocal(0): oText.Y = pLocal(1)
                    Call changeTextProperties(oText)
                    bModified = True
                Case 5 'Top
                    If oPointCol.Count > 1 Then
                        Set oRefText = oTextCollection.Item(2)
                        p1(0) = oPointCol.Item(2).Item("ShtP1")(0): p1(1) = oPointCol.Item(2).Item("ShtP1")(1)
                        p2(0) = oPointCol.Item(2).Item("ShtP2")(0): p2(1) = oPointCol.Item(2).Item("ShtP2")(1)
                        pText(0) = p1(0) + (p2(0) - p1(0)) / 2
                        pText(1) = p1(1) + dDWGTextGAP
                        pLocal = dGetLocalPositionfromGlobal(pText, oText)
                        oText.AnchorPosition = catBottomCenter
                        oText.X = pLocal(0): oText.Y = pLocal(1)
                        Call changeTextProperties(oText)
                        bModified = True
                    End If
                Case 6 'Right
                    If oPointCol.Count > 1 Then
                        Set oRefText = oTextCollection.Item(2)
                        p2(0) = oPointCol.Item(2).Item(2)(0): p2(1) = oPointCol.Item(2).Item(2)(1)
                        p3(0) = oPointCol.Item(2).Item(3)(0): p3(1) = oPointCol.Item(2).Item(3)(1)
                        pText(0) = p2(0) + dDWGTextGAP + 5
                        pText(1) = p3(1) + (p2(1) - p3(1)) / 2
                        pLocal = dGetLocalPositionfromGlobal(pText, oText)
                        oText.AnchorPosition = catMiddleLeft
                        oText.X = pLocal(0): oText.Y = pLocal(1)
                        Call changeTextProperties(oText)
                        bModified = True
                    End If
                Case 7 'Bottom
                    If oPointCol.Count > 1 Then
                        Set oRefText = oTextCollection.Item(2)
                        p3(0) = oPointCol.Item(2).Item(3)(0): p3(1) = oPointCol.Item(2).Item(3)(1)
                        p4(0) = oPointCol.Item(2).Item(4)(0): p4(1) = oPointCol.Item(2).Item(4)(1)
                        pText(0) = p4(0) + (p3(0) - p4(0)) / 2
                        pText(1) = p3(1) - dDWGTextGAP - 2
                        pLocal = dGetLocalPositionfromGlobal(pText, oText)
                        oText.AnchorPosition = catTopCenter
                        oText.X = pLocal(0): oText.Y = pLocal(1)
                        Call changeTextProperties(oText)
                        bModified = True
                    End If
                Case 8 'Left
                    If oPointCol.Count > 1 Then
                        Set oRefText = oTextCollection.Item(2)
                        p1(0) = oPointCol.Item(2).Item(1)(0): p1(1) = oPointCol.Item(2).Item(1)(1)
                        p4(0) = oPointCol.Item(2).Item(4)(0): p4(1) = oPointCol.Item(2).Item(4)(1)
                        pText(0) = p1(0) - dDWGTextGAP - 5
                        pText(1) = p4(1) + (p1(1) - p4(1)) / 2
                        pLocal = dGetLocalPositionfromGlobal(pText, oText)
                        oText.AnchorPosition = catMiddleRight
                        oText.X = pLocal(0): oText.Y = pLocal(1)
                        Call changeTextProperties(oText)
                        bModified = True
                    End If
            End Select
            Call UpdateZONEOfExistingLocationText(oCurrentItem)
            Call UpdateZONEOfExistingLocationText(oAssociatedView)
            '*** Modify value of Vue texte since location of box has changed
            Call UpdateLocationTextsValueInDrawing(oCurrentItem, oAssociatedView)
            Call ChangeExcelValue(oCurrentItem, oAssociatedView, iRow)
        Else
            Exit Sub
        End If
    ElseIf iCol >= 7 Then
        '*** Position Existinglocation text
        If bExistInCol(oAssociatedView, "EXISTINGLOCATIONTXT") Then
            Set oRefText = oAssociatedView.Item("TEXTOBJECT").Item(1)
            Set oExistingLocation = oAssociatedView.Item("EXISTINGLOCATIONTXT")
            Set oTextCollection = oAssociatedView.Item("TEXTOBJECT")
            Set oText = oExistingLocation.Item("TEXTOBJECT").Item(1)
            Set oPointCol = oAssociatedView.Item("POINTCOLLECTION")
            p1(0) = oPointCol.Item(1).Item(1)(0): p1(1) = oPointCol.Item(1).Item(1)(1)
            p2(0) = oPointCol.Item(1).Item(2)(0): p2(1) = oPointCol.Item(1).Item(2)(1)
            pText(0) = p1(0) + (p2(0) - p1(0)) / 2
            pText(1) = p1(1) - dTitleLocHeightAdj
            pText(0) = dXcoodforViewTitleLocationTxt(pText(0), oAssociatedView.Item(1)) '***Right Side oF viewName for View Title
            pLocal = dGetLocalPositionfromGlobal(pText, oRefText)
            oText.AnchorPosition = catBottomLeft
            oText.X = pLocal(0): oText.Y = pLocal(1)
            
            Call UpdateZONEOfExistingLocationText(oAssociatedView)
            Call UpdateZONEOfExistingLocationText(oCurrentItem)
            '*** Modify value of callout texte since location of vuebox has changed
            Call UpdateLocationTextsValueInDrawing(oCurrentItem, oAssociatedView)
            Call ChangeExcelValue(oCurrentItem, oAssociatedView, iRow)
        Else
            Exit Sub
        End If
    End If
    Dim sAr As Variant
    '***update Iterative position of location text --- 1 of 8 or 1 of 4
    Call ModifyExcelSheet(True)
    If bModified Then
        If UBound(Split(oWorksheet.Cells(iRow, 1), ";")) = 2 Then
           sAr = Split(oWorksheet.Cells(iRow, 1), ";")
           sAr(2) = CStr(iLocation)
           oWorksheet.Cells(iRow, 1) = Join(sAr, ";")
        ElseIf UBound(Split(oWorksheet.Cells(iRow, 1), ";")) = 1 Then
            oWorksheet.Cells(iRow, 1) = oWorksheet.Cells(iRow, 1) & ";" & CStr(CStr(iLocation))
        End If
    End If
    Call ModifyExcelSheet(False)
End Sub
Private Sub UpdateLocationTextsValueInDrawing(ByRef oCurrentItem As Collection, ByRef oAssociatedView As Collection)
    Dim oCalloutText As DrawingText
    Dim oVueText As DrawingText
    Dim oExistingCol As Collection
    
    If bExistInCol(oCurrentItem, "EXISTINGLOCATIONTXT") Then
        Set oExistingCol = oCurrentItem.Item("EXISTINGLOCATIONTXT")
        On Error Resume Next
             Set oCalloutText = oExistingCol.Item("TEXTOBJECT").Item(1)
             lerr = Err.Number
         On Error GoTo 0
         If lerr = 0 Then
             If oCalloutText.Text <> Replace(oAssociatedView.Item("ZONE"), "-", vbLf) Then
                 oCalloutText.Text = Replace(oAssociatedView.Item("ZONE"), "-", vbLf)
                 oExistingCol.Remove "ZONE"
                 oExistingCol.Add Replace(oCalloutText.Text, vbLf, "-"), "ZONE"
             End If
             Call changeTextProperties(oCalloutText)
         End If
    End If
    If bExistInCol(oAssociatedView, "EXISTINGLOCATIONTXT") Then
        Set oExistingCol = oAssociatedView.Item("EXISTINGLOCATIONTXT")
        On Error Resume Next
            Set oVueText = oExistingCol.Item("TEXTOBJECT").Item(1)
        On Error GoTo 0
        If lerr = 0 Then
            If oVueText.Text <> Replace(oCurrentItem.Item("ZONE"), "-", vbLf) Then
                 oVueText.Text = Replace(oCurrentItem.Item("ZONE"), "-", vbLf)
                 oExistingCol.Remove "ZONE"
                 oExistingCol.Add Replace(oVueText.Text, vbLf, "-"), "ZONE"
            End If
            Call changeTextProperties(oVueText)
        End If
    End If
End Sub
'********************************************************************************
'* Name: UserForm QueryClose
'* Purpose: When user closes form (pressing X in the corner), close form
'********************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Call cmdCancel_Click
End Sub

Private Sub cmdSelect_Click()
    Dim iRow As Long, iCol As Long, iter As Integer
    Dim iCase As Integer
    Dim oSht As DrawingSheet
    Dim oColl As Collection
    Dim oText As DrawingText
    Dim dViewPt(1) As Double
    Dim dZoom As Double
    Dim oCurrentItem As Collection, oAssociatedView As Collection
    
    If Not sSourceWindow Is Nothing And Not CATIA.ActiveWindow Is sSourceWindow Then sSourceWindow.Activate
    
    On Error Resume Next
        iRow = oExcel.App.ActiveCell.Row
        iCol = oExcel.App.ActiveCell.Column
        lerr = Err.Number
    On Error GoTo 0
    If lerr <> 0 Then Exit Sub  'Cannot do if we cannot get excel object
    
    iCase = 0
    If oWorksheet.Cells(iRow, 1) Like "ListOfCallout;*" Then
        Set oColl = oListOfCallouts
        iCase = 1
    ElseIf oWorksheet.Cells(iRow, 1) Like "ListOfViewsInfo;*" Then
        Set oColl = oListOfVueInfo
        iCase = 2
    ElseIf oWorksheet.Cells(iRow, 1) Like "ListofLocationInfo;*" Then
        Set oColl = oListofLocationInfo
        iCase = 3
    Else
        Exit Sub
    End If
   
    If Not IsNumeric(Split(oWorksheet.Cells(iRow, 1), ";")(UBound(Split(oWorksheet.Cells(iRow, 1), ";")))) Then Exit Sub
    iter = CInt(Split(oWorksheet.Cells(iRow, 1), ";")(1))
    
    Set oCurrentItem = oColl.Item(iter)
    If iCol >= 7 Then
        If bExistInCol(oCurrentItem, "ASSOCIATEDVIEW") Then
            Set oCurrentItem = oCurrentItem.Item("ASSOCIATEDVIEW")
        End If
    End If
    Set oSht = oCurrentItem.Item("SHEET")
    oSht.Activate
    
    Set oText = oCurrentItem.Item("TEXTOBJECT").Item(1)
    CATIA.ActiveDocument.Selection.Clear
    If oCurrentItem.Item("TEXTOBJECT").Count = 1 Then   ' case for detail callout, location info, section view name
        If bExistInCol(oCurrentItem, "EXISTINGLOCATIONTXT") Then
            Call ViewPtAndZoom(oCurrentItem.Item("EXISTINGLOCATIONTXT").Item("POINTCOLLECTION"), dViewPt, dZoom)
        Else
            Call ViewPtAndZoom(oCurrentItem.Item("POINTCOLLECTION"), dViewPt, dZoom)
        End If
    ElseIf oCurrentItem.Item("TEXTOBJECT").Count > 1 Then 'case for section callout
        Call ViewPtAndZoom(oCurrentItem.Item("POINTCOLLECTION"), dViewPt, dZoom)
    End If
    On Error Resume Next
    CATIA.ActiveWindow.ActiveViewer.Viewpoint2D.PutOrigin Array(dViewPt(0), dViewPt(1))
    CATIA.ActiveWindow.ActiveViewer.Viewpoint2D.Zoom = dZoom
    CATIA.ActiveDocument.Selection.Add oCurrentItem.Item("VIEW")
    CATIA.ActiveDocument.Selection.Search "(Drafting.Callout.Visibility=Visible),sel"
    
    CATIA.ActiveDocument.Selection.Search "(Drafting.Text.'Text String'=" & oText.Text & " & Drafting.Text.Visibility=Visible),sel"
    If Err.Number <> 0 Then CATIA.ActiveDocument.Selection.Clear
    On Error GoTo 0
End Sub
'*** To find viewpoint coordinate and zooom level
Private Sub ViewPtAndZoom(ByRef oPointCollection As Collection, ByRef dViewPt() As Double, ByRef dZoom As Double)
    Dim oVct(1) As Double
    Dim dLength As Double
    If oPointCollection.Count = 1 Then 'case for detail
        dViewPt(0) = oPointCollection.Item(1).Item(1)(0)
        dViewPt(1) = oPointCollection.Item(1).Item(1)(1)
        dZoom = 0.01
    ElseIf oPointCollection.Count = 2 Then  'case for section
        oVct(0) = oPointCollection.Item(2).Item(3)(0) - oPointCollection.Item(1).Item(1)(0)
        oVct(1) = oPointCollection.Item(2).Item(3)(1) - oPointCollection.Item(1).Item(1)(1)
        dViewPt(0) = oPointCollection.Item(1).Item(1)(0) + oVct(0) / 2
        dViewPt(1) = oPointCollection.Item(1).Item(1)(1) + oVct(1) / 2
        dZoom = 1.65 * LengthVector2D(oVct) ^ (-0.998)  '*** emperical equation
    End If
End Sub

Private Sub cmdCreate_Click()
    Dim iRow As Long, iCol As Long, iter As Integer
    Dim iCase As Integer
    
    Dim oSht As DrawingSheet
    Dim oColl As Collection
    Dim oText As DrawingText
    Dim oVue As DrawingView
    Dim oRefText As DrawingText
    Dim oCurrentItem As Collection, oAssociatedView As Collection
    Dim dTextwidth As Double, dTextHeight As Double
    Dim dSheetX As Double, dSheetY As Double
    Dim lerr As Long
    
    Dim pText(1) As Double, p1(1) As Double, p2(1) As Double, p4(1) As Double
    Dim pLocal As Variant
    Dim sText As String
    
    Call cmdSelect_Click
    
    On Error Resume Next
        iRow = oExcel.App.ActiveCell.Row
        iCol = oExcel.App.ActiveCell.Column
        lerr = Err.Number
    On Error GoTo 0
    If lerr <> 0 Then Exit Sub  'Cannot do if we cannot get excel object
    
    
    iCase = 0
    If oWorksheet.Cells(iRow, 1) Like "ListOfCallout;*" Then    '***Can create only if Firt column has value "ListOfCallout" otherwise we have insufficient info
        Set oColl = oListOfCallouts
        iCase = 1
    Else
        Exit Sub
    End If
    
    If Not IsNumeric(Split(oWorksheet.Cells(iRow, 1), ";")(UBound(Split(oWorksheet.Cells(iRow, 1), ";")))) Then Exit Sub
    iter = CInt(Split(oWorksheet.Cells(iRow, 1), ";")(1))
    Set oCurrentItem = oColl.Item(iter)

    Set oAssociatedView = oCurrentItem.Item("ASSOCIATEDVIEW")   '*** This should exist in List of callout
    If iCol < 7 Then    '****Create Text around Callout
        If bExistInCol(oCurrentItem, "EXISTINGLOCATIONTXT") Then
            Call UpdateLocationTextsValueInDrawing(oCurrentItem, oAssociatedView)
        Else
            '**** To ensure that user has not deleted the text
            On Error Resume Next
            Set oRefText = oCurrentItem.Item("TEXTOBJECT").Item(1)
            lerr = Err.Number
            On Error GoTo 0
            If lerr <> 0 Then Exit Sub  'Cannot do anything if text deleted
            Set oVue = oRefText.Parent.Parent
            Set oSht = oVue.Parent.Parent
            oSht.Activate
            '**** Default location for text creation

            p1(0) = oCurrentItem.Item("POINTCOLLECTION").Item(1).Item(1)(0): p1(1) = oCurrentItem.Item("POINTCOLLECTION").Item(1).Item(1)(1)
            p2(0) = oCurrentItem.Item("POINTCOLLECTION").Item(1).Item(2)(0): p2(1) = oCurrentItem.Item("POINTCOLLECTION").Item(1).Item(2)(1)
            pText(0) = p1(0) + (p2(0) - p1(0)) / 2
            pText(1) = p1(1) + dDWGTextGAP
            pLocal = dGetLocalPositionfromGlobal(pText, oRefText)
            sText = Replace(oAssociatedView.Item("ZONE"), "-", vbLf)
            Set oText = oVue.Texts.Add(Replace(oAssociatedView.Item("ZONE"), "-", vbLf), pLocal(0), pLocal(1))
            oText.AnchorPosition = catBottomCenter
            Call changeTextProperties(oText)
            '***create "EXISTINGLOCATIONTXT" collection and populate
            Call CreateExistingLocation(oCurrentItem, oText)
            Call UpdateLocationTextsValueInDrawing(oCurrentItem, oAssociatedView)
        End If
    ElseIf iCol >= 7 Then   '*** Create Text around View Name
        If bExistInCol(oCurrentItem, "EXISTINGLOCATIONTXT") Then
            If bExistInCol(oAssociatedView, "EXISTINGLOCATIONTXT") Then
                 Call UpdateLocationTextsValueInDrawing(oCurrentItem, oAssociatedView)
            Else
                On Error Resume Next
                Set oRefText = oAssociatedView.Item("TEXTOBJECT").Item(1)
                lerr = Err.Number
                On Error GoTo 0
                If lerr <> 0 Then Exit Sub  'Cannot do anything if text deleted
                Set oVue = oRefText.Parent.Parent
                Set oSht = oVue.Parent.Parent
                oSht.Activate
                p1(0) = oAssociatedView.Item("POINTCOLLECTION").Item(1).Item(1)(0): p1(1) = oAssociatedView.Item("POINTCOLLECTION").Item(1).Item(1)(1)
                p2(0) = oAssociatedView.Item("POINTCOLLECTION").Item(1).Item(2)(0): p2(1) = oAssociatedView.Item("POINTCOLLECTION").Item(1).Item(2)(1)
                pText(0) = p1(0) + (p2(0) - p1(0)) / 2
                pText(1) = p1(1) - dTitleLocHeightAdj
                pText(0) = dXcoodforViewTitleLocationTxt(pText(0), oAssociatedView.Item(1)) '***Right Side oF viewName for View Title
                pLocal = dGetLocalPositionfromGlobal(pText, oRefText)
                'sText = Replace(oAssociatedView.Item("ZONE"), "-", vbLf)
                Set oText = oVue.Texts.Add(Replace(oCurrentItem.Item("ZONE"), "-", vbLf), pLocal(0), pLocal(1))
                oText.AnchorPosition = catBottomLeft
                Call changeTextProperties(oText)
                '***create "EXISTINGLOCATIONTXT" collection and populate
                Call CreateExistingLocation(oAssociatedView, oText)
                Call UpdateLocationTextsValueInDrawing(oCurrentItem, oAssociatedView)
            End If
        Else
            MsgBox "First Create LocationText for callout", vbCritical
        End If
    End If
    '***Update Excel file to latest values
    Call ChangeExcelValue(oCurrentItem, oAssociatedView, iRow)
End Sub
'***Change Accpeted and existing zone values of current row
Private Sub ChangeExcelValue(ByVal oCurrentItem As Collection, ByVal oAssociatedView As Collection, ByVal iRow As Long)
    Call ModifyExcelSheet(True)
    oWorksheet.Cells(iRow, iACCPLCNVUE) = oCurrentItem.Item("ZONE")
    oWorksheet.Cells(iRow, iACCPLCNCALL) = oAssociatedView.Item("ZONE")
    If bExistInCol(oCurrentItem, "EXISTINGLOCATIONTXT") Then
        oWorksheet.Cells(iRow, iEXSTLCNCALL) = oCurrentItem.Item("EXISTINGLOCATIONTXT").Item("ZONE")
    End If
    If bExistInCol(oAssociatedView, "EXISTINGLOCATIONTXT") Then
        oWorksheet.Cells(iRow, iEXSTLCNVUE) = oAssociatedView.Item("EXISTINGLOCATIONTXT").Item("ZONE")
    End If
    Call ModifyExcelSheet(False)
    Call VFNCheckExcelRow(iRow, oWorksheet)
End Sub

Private Sub CreateExistingLocation(ByRef oCurrentItem As Collection, ByVal oText As DrawingText)
    Dim p1(1) As Double, p2(1) As Double, p4(1) As Double
    Dim dSheetX As Double, dSheetY As Double, dTextHeight As Double, dTextwidth As Double
    
    '***create "EXISTINGLOCATIONTXT" collection and populate
    oCurrentItem.Add New Collection, "EXISTINGLOCATIONTXT"
    
    oCurrentItem.Item("EXISTINGLOCATIONTXT").Add Join(Split(oText.Text, vbLf), "-"), "ZONE"
    oCurrentItem.Item("EXISTINGLOCATIONTXT").Add New Collection, "TEXTOBJECT"
    oCurrentItem.Item("EXISTINGLOCATIONTXT").Item("TEXTOBJECT").Add oText
    dTextwidth = dDrawingTextwidth(oText, dTextHeight)
    oCurrentItem.Item("EXISTINGLOCATIONTXT").Add New Collection, "POINTCOLLECTION"
    oCurrentItem.Item("EXISTINGLOCATIONTXT").Item("POINTCOLLECTION").Add oTextBoundingBox(oText, dTextwidth, dTextHeight), "POINTCOLLECTION"
    oCurrentItem.Item("EXISTINGLOCATIONTXT").Add oText.Parent.Parent, "VIEW"
    oCurrentItem.Item("EXISTINGLOCATIONTXT").Add oText.Parent.Parent.Parent.Parent, "SHEET"
    
    Call UpdateZONEOfExistingLocationText(oCurrentItem)
End Sub
'*** Update zone info as per existing location
Private Sub UpdateZONEOfExistingLocationText(ByRef oItem As Collection)
    Dim p1(1) As Double, p2(1) As Double, p4(1) As Double
    Dim dSheetX As Double, dSheetY As Double, dTextHeight As Double, dTextwidth As Double
    Dim oExistingLocation As Collection
    Dim oText As DrawingText
    
    If bExistInCol(oItem, "EXISTINGLOCATIONTXT") Then
        Set oExistingLocation = oItem.Item("EXISTINGLOCATIONTXT")
        oExistingLocation.Item("POINTCOLLECTION").Remove 1
        Set oText = oExistingLocation.Item("TEXTOBJECT").Item(1)
        dTextwidth = dDrawingTextwidth(oText, dTextHeight)
        oExistingLocation.Item("POINTCOLLECTION").Add oTextBoundingBox(oText, dTextwidth, dTextHeight), "POINTCOLLECTION"
        p1(0) = oExistingLocation.Item("POINTCOLLECTION").Item(1).Item(1)(0): p1(1) = oExistingLocation.Item("POINTCOLLECTION").Item(1).Item(1)(1)
        p2(0) = oExistingLocation.Item("POINTCOLLECTION").Item(1).Item(2)(0)
        p4(1) = oExistingLocation.Item("POINTCOLLECTION").Item(1).Item(4)(1)
        
        '***Find center of location info box
        dSheetX = p1(0) + (p2(0) - p1(0)) / 2
        dSheetY = p4(1) + (p1(1) - p4(1)) / 2
        '***Set "ZONE" value to zone of Existing Text
        oExistingLocation.Remove "ZONE"
        oExistingLocation.Add Replace(oText.Text, vbLf, "-"), "ZONE"
        
        oItem.Remove "ZONE"
        oItem.Add sZoneOfTextInDrawing(oText, dSheetX, dSheetY), "ZONE"
    End If

End Sub
Public Sub ModifyExcelSheet(ByVal bLock As Boolean)
'***Lock the sheet
oExcel.App.DisplayAlerts = False
If Not bLock Then
    On Error Resume Next
    oWorksheet.Protect sEXCELPWD, True, True
    On Error GoTo 0
Else
    On Error Resume Next
    oWorksheet.Unprotect sEXCELPWD
    On Error GoTo 0
End If
oExcel.App.DisplayAlerts = True
End Sub

Private Sub changeTextProperties(ByRef oText As DrawingText)
    Dim oSelcn As Selection
    Dim i As Integer
    Dim oVisProperty As VisPropertySet
    
    Set oSelcn = CATIA.ActiveDocument.Selection
    
    oText.TextProperties.Justification = catCenter
    oText.TextProperties.FontName = "SSS1"
    oText.TextProperties.FontSize = 4.064
    oText.FrameType = 52
    
    If Me.chkPink.Value Then
        oSelcn.Clear
        oSelcn.Add oText.Parent.Parent
        oSelcn.Search "(Drafting.Text.'Text String'=*),sel"
        For i = oSelcn.Count To 1 Step -1
            If Not (oSelcn.Item(i).Value.Text = oText.Text And oSelcn.Item(i).Value.FrameType = 52) Then oSelcn.Remove2 i
        Next
        Set oVisProperty = oSelcn.VisProperties
        Call oVisProperty.SetRealColor(255, 0, 255, False)
        oSelcn.Clear
    ElseIf Me.chkGreen.Value Then
        oSelcn.Clear
        oSelcn.Add oText.Parent.Parent
        oSelcn.Search "(Drafting.Text.'Text String'=*),sel"
        For i = oSelcn.Count To 1 Step -1
            If Not (oSelcn.Item(i).Value.Text = oText.Text And oSelcn.Item(i).Value.FrameType = 52) Then oSelcn.Remove2 i
        Next
        Set oVisProperty = oSelcn.VisProperties
        Call oVisProperty.SetRealColor(0, 128, 0, False)
        oSelcn.Clear
    End If
End Sub
Private Function dGetLocalPositionfromGlobal(ByRef dGlobal() As Double, ByVal oText As DrawingText) As Variant
    Dim oDrawingView As DrawingView
    Dim oMainVue As DrawingView
    Dim dSheetAngle As Double
    Dim oFactory2d As Factory2D
    Dim dSheetX As Double
    Dim dSheetY As Double
    Dim dViewY As Double
    Dim dViewScale As Double
    Dim dViewAngle As Double
    Dim dViewPt(1) As Double
    Set oDrawingView = oText.Parent.Parent
    dViewScale = oDrawingView.Scale
    dViewAngle = oDrawingView.Angle
    dSheetX = -oDrawingView.xAxisData
    dSheetY = -oDrawingView.yAxisData

    dViewPt(0) = ((dSheetX * Cos(dViewAngle) + dSheetY * Sin(dViewAngle)) + (dGlobal(0) * Cos(dViewAngle) + dGlobal(1) * Sin(dViewAngle))) / dViewScale
    dViewPt(1) = ((-dSheetX * Sin(dViewAngle) + dSheetY * Cos(dViewAngle)) + (-dGlobal(0) * Sin(dViewAngle) + dGlobal(1) * Cos(dViewAngle))) / dViewScale
    dGetLocalPositionfromGlobal = dViewPt
End Function

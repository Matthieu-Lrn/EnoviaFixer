Attribute VB_Name = "mdlVueFindNumber"
Private Const MAXGAPCallout As Double = 15.24
Private Const MAXGAPView As Double = 25.4
Public Const iEXSTLCNCALL As Integer = 4
Public Const iACCPLCNCALL As Integer = 5
Public Const iSTATUSCALL As Integer = 6
Public Const iEXSTLCNVUE As Integer = 8
Public Const iACCPLCNVUE As Integer = 9
Public Const iSTATUSVUE As Integer = 10
Public Const sEXCELPWD As String = "AutomationLock"
Public Const dTitleLocHeightAdj As Double = 6.53
Public Const dDWGTextGAP As Double = 2.54
Private oDrawingDocument As DrawingDocument
Private sExcelReportTemplate As String
Private sSourceWindow As Window
'********************************************************************************
'* Name: Vue Location Info Tool
'* Purpose: Check/Create /Update Location info text in a drawing for Detail and Section view
'*
'* Assumption:
'*
'* Author: Abhishek Kamboj
'* Updated by:
'* Language: VBA
'********************************************************************************
Public Sub LaunchVueFindNumber(ByVal sdummy As String)

    Dim oDummyColl As Collection
    Dim oSelcn As Selection
    Dim oSelectedItm As SelectedElement
    Dim oListOCallouts As Collection
    Dim oListOfViewsInfo As Collection
    Dim oCombinedInfo As Collection
    Dim oListofLocationInfo As Collection
    
    Dim oSheet As DrawingSheet
    Dim oDrawingView As DrawingView
    Dim iViewNumber As Integer
    Dim iTextNumber As Integer
    Dim iSheetNb As Integer
    Dim sSheetNb As String
    
    Dim dViewX As Double
    Dim dViewY As Double
    Dim dViewScale As Double
    Dim dViewAngle As Double
    
    Dim oText As DrawingText
    Dim dTextX As Double
    Dim dTextY As Double
    
    Dim dSheetCoordX As Double
    Dim dSheetCoordY As Double
    
    Dim sZone As String
    
    Dim sString As String
    Dim sWarningPageSize As String
    Dim sWarningOutsideFrame As String
    
    Dim bMultiSheet As Boolean
    
    Dim oDwgSheetsFindNo As New clsCollection
    Dim oItem As clsCollection
    
    Dim dFontsize As Double
    
    Dim oTextFct As New clsDwgTextFunctions
    Dim dPaperWidth As Double
    Dim dPaperHeight As Double
    Dim iNbHorZone As Integer
    Dim iNbVerZone As Integer
    Dim sPageSize As String
    Dim sSearchstring As String
    Dim dTextwidth As Double
    Dim dTextHeight As Double
    Dim lerr As Long
    Dim oSecCol As Collection
    Dim bResetTB As Boolean

    '***Progress Form
    Call frmProgress.progressBarInitialize("Searching for " & sSearchObject)
    '***On retrouve le drawing document actif
    Set sSourceWindow = CATIA.ActiveWindow
    Set oDrawingDocument = CATIA.ActiveDocument
    '***Log File
    Call AddToLogFile("Vue Location Info", oDrawingDocument.Name)
    '***
    Set oSelcn = oDrawingDocument.Selection
    '***We check if we have more than one SH sheet
    bMultiSheet = oTextFct.CheckMultiSheet(oDrawingDocument)
    '***Initialize
    sWarningPageSize = ""
    sWarningOutsideFrame = ""
    Set oListOCallouts = New Collection
    Set oListOfViewsInfo = New Collection
    Set oListofLocationInfo = New Collection
    '***---------Get Template path---------
    sExcelReportTemplate = oPathDict("KBE_LocationInfoReport")
    If sExcelReportTemplate = "" Then
        MsgBox "Execution aborted because tool cannot Find Report Template." & vbCrLf & "Please contact KBE team.", vbExclamation
        bResetTB = True
        GoTo endsub
    End If
    '***----------------------------------------------------BEGIN Scanning the Drawing
    For iSheetNb = 1 To oDrawingDocument.Sheets.Count
        '***Sheet info
        Set oSheet = oDrawingDocument.Sheets.Item(iSheetNb)
        dPaperWidth = oSheet.GetPaperWidth
        dPaperHeight = oSheet.GetPaperHeight
        Call oTextFct.FindPaperSize(dPaperWidth, dPaperHeight, sPageSize, iNbVerZone, iNbHorZone)
        '***Check Page size
        If sPageSize = "Not a Bombardier Format" Then
            sString = " - " & oSheet.Name
            If sWarningPageSize = "" Then
                sWarningPageSize = sString
            Else
                sWarningPageSize = sWarningPageSize & vbCrLf & sString
            End If
            GoTo TheNextSheet
        End If
        '***If sheet name starts with "SH"
        If Left(oSheet.Name, 2) = "SH" Then
            '***Find all Location text in the sheet
            oSelcn.Clear
            oSelcn.Add oSheet
            oSelcn.Search "(Drafting.Text.Visibility=Visible),sel"
            Call frmProgress.progressBarRepaint("Scanning in: " & oSheet.Name)
            For i = 1 To oSelcn.Count
                On Error Resume Next
                If bIsLocationInfoText(oSelcn.Item(i).Value) Then
                    oListofLocationInfo.Add New Collection, CStr(oListofLocationInfo.Count + 1)
                    oListofLocationInfo.Item(oListofLocationInfo.Count).Add CStr(oListofLocationInfo.Count)
                    Set oText = oSelcn.Item(i).Value
                    oListofLocationInfo.Item(oListofLocationInfo.Count).Add Join(Split(oText.Text, vbLf), "-"), "ZONE"
                    oListofLocationInfo.Item(oListofLocationInfo.Count).Add New Collection, "TEXTOBJECT"
                    oListofLocationInfo.Item(oListofLocationInfo.Count).Item("TEXTOBJECT").Add oSelcn.Item(i).Value
                    dTextwidth = dDrawingTextwidth(oText, dTextHeight)
                    oListofLocationInfo.Item(oListofLocationInfo.Count).Add New Collection, "POINTCOLLECTION"
                    oListofLocationInfo.Item(oListofLocationInfo.Count).Item("POINTCOLLECTION").Add oTextBoundingBox(oText, dTextwidth, dTextHeight), "POINTCOLLECTION"
                    oListofLocationInfo.Item(oListofLocationInfo.Count).Add oText.Parent.Parent, "VIEW"
                    oListofLocationInfo.Item(oListofLocationInfo.Count).Add oSheet, "SHEET"
                End If
                On Error GoTo 0
            Next
            'Looping all views
            For iViewNumber = 1 To oSheet.Views.Count
                Set oDrawingView = oSheet.Views.Item(iViewNumber)
                '*** Progress form
                Call frmProgress.progressBarRepaint("Scanning in: " & oSheet.Name & " / " & oDrawingView.Name)
                If bCancelAction = True Or GetAsyncKeyState(vbKeyEscape) <> 0 Then
                    bResetTB = True
                    GoTo endsub
                End If
                '*** Get View Name text for section and detail view
                Dim sViewNamestr As String
                Dim sNameForViewinfoList As String
                Dim sDisplayName As String
                sSearchstring = ""
                sNameForViewinfoList = ""
                sViewNamestr = sNoMultiSpace(oDrawingView.Name)
                '*** Searching for following patterns : SECTION A-A or VIEW A-A or DETAIL A or VIEW A
                If sViewNamestr Like "SECTION *" Then
                    If sViewNamestr Like "SECTION [A-Z]-[A-Z]" Or sViewNamestr Like "SECTION [A-Z]-[A-Z] *" Then
                        sSearchstring = Left(sViewNamestr, 11)
                        sNameForViewinfoList = sSearchstring
                    ElseIf sViewNamestr Like "SECTION [A-Z][A-Z]-[A-Z][A-Z]" Or sViewNamestr Like "SECTION [A-Z][A-Z]-[A-Z][A-Z] *" Then
                        sSearchstring = Left(sViewNamestr, 13)
                        sNameForViewinfoList = sSearchstring
                    End If
                ElseIf sViewNamestr Like "DETAIL *" Then
                    If sViewNamestr Like "DETAIL [A-Z][A-Z]" Or sViewNamestr Like "DETAIL [A-Z][A-Z] *" Then
                        sSearchstring = Left(sViewNamestr, 9)
                        sNameForViewinfoList = sSearchstring
                    ElseIf sViewNamestr Like "DETAIL [A-Z]" Or sViewNamestr Like "DETAIL [A-Z] *" Then
                        sSearchstring = Left(sViewNamestr, 8)
                        sNameForViewinfoList = sSearchstring
                    End If
                ElseIf sViewNamestr Like "VIEW *" Then 'Some stupid guys rename section views as view ... to cover that
                    If sViewNamestr Like "VIEW [A-Z][A-Z]-[A-Z][A-Z]" Or sViewNamestr Like "VIEW [A-Z][A-Z]-[A-Z][A-Z] *" Then
                        sSearchstring = Left(sViewNamestr, 10)
                        sNameForViewinfoList = "SECTION " & Mid(sSearchstring, 6, 5)
                    ElseIf sViewNamestr Like "VIEW [A-Z]-[A-Z]" Or sViewNamestr Like "VIEW [A-Z]-[A-Z] *" Then
                        sSearchstring = Left(sViewNamestr, 8)
                        sNameForViewinfoList = "SECTION " & Mid(sSearchstring, 6, 3)
                    End If
                End If
                lerr = 0
                On Error Resume Next
                Dim oDummy As Collection
                Set oDummy = oListOfViewsInfo.Item(sSearchstring)
                lerr = Err.Number   '*** if lerr <>0 then item doesn't exist in list
                On Error GoTo 0
                If sSearchstring <> "" And lerr <> 0 Then
                    oSelcn.Clear
                    oSelcn.Add oDrawingView
                    oSelcn.Search "(Drafting.Text.'Text String'='" & sSearchstring & "'*),sel"
                    If oSelcn.Count = 1 Then    '*** View is added only if name of the view found in text
                        Set oText = oSelcn.Item(1).Value
                        dTextwidth = dDrawingTextwidth(oText, dTextHeight)
                        Set oDummyColl = oTextBoundingBox(oText, dTextwidth, dTextHeight)
                        dSheetCoordX = oDummyColl.Item(1)(0) + (oDummyColl.Item(2)(0) - oDummyColl.Item(1)(0)) / 2
                        dSheetCoordX = dXcoodforViewTitleLocationTxt(dSheetCoordX, sSearchstring)
                        dSheetCoordY = oDummyColl.Item(1)(1) - dTitleLocHeightAdj
                        sZone = sZoneOfTextInDrawing(oText, dSheetCoordX, dSheetCoordY)
                        If Not bExistInCol(oListOfViewsInfo, sNameForViewinfoList) Then   '***No Duplicates
                            oListOfViewsInfo.Add New Collection, sNameForViewinfoList
                            oListOfViewsInfo.Item(oListOfViewsInfo.Count).Add sNameForViewinfoList
                            oListOfViewsInfo.Item(oListOfViewsInfo.Count).Add New Collection, "TEXTOBJECT"
                            oListOfViewsInfo.Item(oListOfViewsInfo.Count).Item("TEXTOBJECT").Add oText
                            oListOfViewsInfo.Item(oListOfViewsInfo.Count).Add sZone, "ZONE"
                            dTextwidth = dDrawingTextwidth(oText, dTextHeight)
                            oListOfViewsInfo.Item(oListOfViewsInfo.Count).Add New Collection, "POINTCOLLECTION"
                            oListOfViewsInfo.Item(oListOfViewsInfo.Count).Item("POINTCOLLECTION").Add oTextBoundingBox(oText, dTextwidth, dTextHeight)
                            oListOfViewsInfo.Item(oListOfViewsInfo.Count).Add oDrawingView, "VIEW"
                            oListOfViewsInfo.Item(oListOfViewsInfo.Count).Add oSheet, "SHEET"
                            oListOfViewsInfo.Item(oListOfViewsInfo.Count).Add sSearchstring, "DISPLAYNAME"
                        End If
                    End If
                End If
                '*** Check for all visible callout text in a view
                Set oDummyColl = New Collection
                Set oSelcn = oDrawingDocument.Selection
                oSelcn.Clear
                oSelcn.Add oDrawingView
                oSelcn.Search "(Drafting.Callout.Visibility=Visible),sel"
                oSelcn.Search "(Drafting.Text.Visibility=Visible),sel"
                
                For i = 1 To oSelcn.Count
                    On Error Resume Next
                    oDummyColl.Add oSelcn.Item(i).Value, oSelcn.Item(i).Value.Text
                    On Error GoTo 0
                Next
                '*** we loop all the callouts
                For iTextNumber = 1 To oDummyColl.Count
                    Set oText = oDummyColl.Item(iTextNumber)
                    Set oSecCol = New Collection
                    oSelcn.Clear
                    oSelcn.Add oDrawingView
                    oSelcn.Search "(Drafting.Callout.Visibility=Visible),sel"
                    oSelcn.Search "(Drafting.Text.'Text String'='" & oText.Text & "' & Drafting.Text.Visibility=Visible),sel"
                    
                    If oSelcn.Count > 1 Then   'Incase of section we have two callouttext
                        For i = 1 To oSelcn.Count
                            oSecCol.Add oSelcn.Item(i).Value
                        Next
                        If oSelcn.Count > 2 Then    '*** flag a message and end execution.
                            MsgBox "Too many callouts of: " & oText.Text & " in " & oDrawingView.Name & " of " & oSheet.Name & vbCrLf & _
                                   "Tool is not compatible with the drawing." & vbCrLf & _
                                   "Callout Letter should be unique for each section and detail view in  the drawing.", vbCritical
                            bResetTB = True
                            GoTo endsub
                        End If
                        For i = oSelcn.Count To 2 Step -1
                            oSelcn.Remove2 i
                        Next
                    ElseIf oSelcn.Count = 1 Then    ' Incase of detail we have only one callout text
                        oSecCol.Add oSelcn.Item(1).Value
                    End If
                    '****get zone
                    For i = 1 To oSecCol.Count
                        If i = 1 Then sZone = sZoneOfTextInDrawing(oSecCol.Item(i)) Else sZone = sZone & ";" & sZoneOfTextInDrawing(oSecCol.Item(i))
                    Next
                    'If otext.Text = "A" Then Stop
                    dTextwidth = dDrawingTextwidth(oText, dTextHeight)
                    If CATIA.StatusBar Like "*Detail View*" Then
                        If Not bExistInCol(oListOCallouts, "DETAIL " & Trim(oText.Text)) Then
                            oListOCallouts.Add New Collection, "DETAIL " & Trim(oText.Text)
                            oListOCallouts.Item(oListOCallouts.Count).Add "DETAIL " & oText.Text
                            oListOCallouts.Item(oListOCallouts.Count).Add New Collection, "TEXTOBJECT"
                            oListOCallouts.Item(oListOCallouts.Count).Item("TEXTOBJECT").Add oText
                            oListOCallouts.Item(oListOCallouts.Count).Add sZone, "ZONE"
                            oListOCallouts.Item(oListOCallouts.Count).Add New Collection, "POINTCOLLECTION"
                            oListOCallouts.Item(oListOCallouts.Count).Item("POINTCOLLECTION").Add oTextBoundingBox(oText, dTextwidth, dTextHeight)
                            oListOCallouts.Item(oListOCallouts.Count).Add oDrawingView, "VIEW"
                            oListOCallouts.Item(oListOCallouts.Count).Add oSheet, "SHEET"
                            oListOCallouts.Item(oListOCallouts.Count).Add "Callout " & Trim(oText.Text), "DISPLAYNAME"
                        End If
                    ElseIf CATIA.StatusBar Like "*Section View*" Then
                        If Not bExistInCol(oListOCallouts, "SECTION " & Trim(oText.Text) & "-" & Trim(oText.Text)) Then
                             oListOCallouts.Add New Collection, "SECTION " & Trim(oText.Text) & "-" & Trim(oText.Text)
                             oListOCallouts.Item(oListOCallouts.Count).Add "SECTION " & oText.Text & "-" & oText.Text
                             oListOCallouts.Item(oListOCallouts.Count).Add New Collection, "TEXTOBJECT"
                             oListOCallouts.Item(oListOCallouts.Count).Add sZone, "ZONE"
                             oListOCallouts.Item(oListOCallouts.Count).Add New Collection, "POINTCOLLECTION"
                             For i = 1 To oSecCol.Count
                                  oListOCallouts.Item(oListOCallouts.Count).Item("TEXTOBJECT").Add oSecCol.Item(i)
                                  oListOCallouts.Item(oListOCallouts.Count).Item("POINTCOLLECTION").Add oTextBoundingBox(oSecCol.Item(i), dTextwidth, dTextHeight)
                             Next
                             oListOCallouts.Item(oListOCallouts.Count).Add oDrawingView, "VIEW"
                             oListOCallouts.Item(oListOCallouts.Count).Add oSheet, "SHEET"
                             oListOCallouts.Item(oListOCallouts.Count).Add "Callout " & Trim(oText.Text), "DISPLAYNAME"
                         End If
                    ElseIf CATIA.StatusBar Like "*Auxiliary View*" Then
                        If Not bExistInCol(oListOCallouts, "SECTION " & Trim(oText.Text) & "-" & Trim(oText.Text)) Then
                             oListOCallouts.Add New Collection, "SECTION " & Trim(oText.Text) & "-" & Trim(oText.Text)
                             oListOCallouts.Item(oListOCallouts.Count).Add "SECTION " & oText.Text & "-" & oText.Text
                             oListOCallouts.Item(oListOCallouts.Count).Add New Collection, "TEXTOBJECT"
                             oListOCallouts.Item(oListOCallouts.Count).Add sZone, "ZONE"
                             oListOCallouts.Item(oListOCallouts.Count).Add New Collection, "POINTCOLLECTION"
                             For i = 1 To oSecCol.Count
                                  oListOCallouts.Item(oListOCallouts.Count).Item("TEXTOBJECT").Add oSecCol.Item(i)
                                  oListOCallouts.Item(oListOCallouts.Count).Item("POINTCOLLECTION").Add oTextBoundingBox(oSecCol.Item(i), dTextwidth, dTextHeight)
                             Next
                             oListOCallouts.Item(oListOCallouts.Count).Add oDrawingView, "VIEW"
                             oListOCallouts.Item(oListOCallouts.Count).Add oSheet, "SHEET"
                             oListOCallouts.Item(oListOCallouts.Count).Add "Callout " & Trim(oText.Text), "DISPLAYNAME"
                         End If
                    
                    End If
                Next
            Next
        End If
TheNextSheet:
    Next
    '***----------------------------------------------------END Scanning the Drawing

    Call FindAssociatedLocationText(oListOCallouts, oListofLocationInfo, MAXGAPCallout)
    Call FindAssociatedLocationText(oListOfViewsInfo, oListofLocationInfo, MAXGAPView)
    '*** Associate Callout with View
    For i = 1 To oListOCallouts.Count
        If bExistInCol(oListOfViewsInfo, oListOCallouts.Item(i).Item(1)) Then
            oListOCallouts.Item(i).Add oListOfViewsInfo.Item(oListOCallouts.Item(i).Item(1)), "ASSOCIATEDVIEW"
            oListOfViewsInfo.Remove oListOCallouts.Item(i).Item(1)
        End If
    Next
    
    Call SendToExcel(oListOCallouts, oListOfViewsInfo, oListofLocationInfo)
endsub:
    frmProgress.Hide
    If bResetTB Then frmKBEMain.resetToolbar
End Sub
'*** Function to identify if a given text is a location text or not
Private Function bIsLocationInfoText(ByVal oText As DrawingText) As Boolean
    If (oText.FrameType = 52 Or oText.FrameType = CatTextFrameType.catSquare) And _
                    (oText.Text Like "*" & vbLf & "*") And _
                    (Len(oText.Text) = 4 Or Len(oText.Text) = 5) And _
                    (oText.Leaders.Count = 0) Then
                    bIsLocationInfoText = True
    Else
        bIsLocationInfoText = False
    End If
End Function
Private Function sNoMultiSpace(ByVal sexpr As String) As String
    Dim oVar As Variant
    Dim i As Integer
    Do While InStr(1, sexpr, "  ")
        sexpr = Replace(sexpr, "  ", " ")
    Loop
    
    If InStr(1, sexpr, "-") Then
        oVar = Split(sexpr, "-")
        For i = 0 To UBound(oVar)
            oVar(i) = Trim(oVar(i))
        Next
        sexpr = Join(oVar, "-")
    End If
    sNoMultiSpace = sexpr
End Function
Public Function bExistInCol(ByVal oCollection As Collection, ByVal sKey As String) As Boolean
    '**** to find if particular key exist in collection or not
    Dim iObj As AnyObject
    Dim oCol As Collection
    Dim jObj As Variant
    Dim lerr As Long
    
    On Error Resume Next
        Set iObj = oCollection.Item(sKey)
        If Err.Number <> 0 Then
            Err.Clear
            Set oCol = oCollection.Item(sKey)
        End If
        If Err.Number <> 0 Then
            Err.Clear
            jObj = oCollection.Item(sKey)
        End If
        If Err.Number <> 0 Then
            Err.Clear
            iObj = oCollection.Item(sKey)
        End If
        lerr = Err.Number
    On Error GoTo 0
    If lerr = 0 Then bExistInCol = True
End Function
Public Function sZoneOfTextInDrawing(ByRef oText As DrawingText, Optional ByVal dSheetCoordX As Double = 0, Optional ByVal dSheetCoordY As Double = 0) As String
    Dim oTextFct As clsDwgTextFunctions
    Dim dPaperWidth As Double, dPaperHeight As Double, dViewX As Double, dViewY As Double, dViewScale As Double, dViewAngle As Double
    Dim dTextX As Double, dTextY As Double
    Dim sPageSize As String
    Dim iNbVerZone As Integer, iNbHorZone As Integer
    Dim oSheet As DrawingSheet
    Dim oDrawingView As DrawingView
    Dim bMultiSheet As Boolean
    
    Set oTextFct = New clsDwgTextFunctions
    Set oDrawingView = oText.Parent.Parent
    Set oSheet = oDrawingView.Parent.Parent
    
    dPaperWidth = oSheet.GetPaperWidth
    dPaperHeight = oSheet.GetPaperHeight
    
    '*** Position of View in the sheet
    dViewX = oDrawingView.xAxisData
    dViewY = oDrawingView.yAxisData
    dViewScale = oDrawingView.Scale
    dViewAngle = oDrawingView.Angle
    
    bMultiSheet = oTextFct.CheckMultiSheet(oSheet.Parent.Parent)
    Call oTextFct.FindPaperSize(dPaperWidth, dPaperHeight, sPageSize, iNbVerZone, iNbHorZone)
    
    dTextX = oText.X
    dTextY = oText.Y
    
    If dSheetCoordX = 0 Then dSheetCoordX = dViewX + (dTextX * Cos(dViewAngle) - dTextY * Sin(dViewAngle)) * dViewScale
    If dSheetCoordY = 0 Then dSheetCoordY = dViewY + (dTextX * Sin(dViewAngle) + dTextY * Cos(dViewAngle)) * dViewScale
    
    sZone = oTextFct.FindZone(sPageSize, dSheetCoordX, dSheetCoordY, oSheet.Name, bMultiSheet, dPaperWidth, dPaperHeight, iNbVerZone, iNbHorZone)
    sZoneOfTextInDrawing = sZone
End Function
Private Sub FindAssociatedLocationText(ByRef ViewOrCalloutCol As Collection, ByRef oListofLocationInfo As Collection, ByVal dMaxgap As Double)
    '*** Find all associated Location text for callouts
    Dim dMindist As Double, dDist1 As Double, dDist2 As Double, dDist3 As Double, dDist4 As Double
    Dim sSheetName As Double
    Dim oLocationBox As Collection
    Dim oLocationCol  As Collection
    Dim oTextBox As Collection
    Dim sLocationColName As String
    Dim oDistCol As Collection
    Dim iObj
    Dim lerr As Long
    Dim i As Integer, j As Integer, k As Integer, iTextNumber As Integer
    Dim p1(1) As Double, p2(1) As Double, p4(1) As Double
    Dim dSheetX As Double, dSheetY As Double
    
    Dim sLocationText
    For i = 1 To ViewOrCalloutCol.Count
        dMindist = dMaxgap
        sLocationColName = ""
        k = 1
        For Each oTextBox In ViewOrCalloutCol.Item(i).Item("POINTCOLLECTION")
            For Each oLocationCol In oListofLocationInfo
                'If oLocationCol.Item(1) = "10" Then Stop
                If ViewOrCalloutCol.Item(i).Item("SHEET").Name = oLocationCol.Item("SHEET").Name Then '***search in same sheet
                    Set oLocationBox = oLocationCol.Item("POINTCOLLECTION").Item(1)
                    Set oDistCol = New Collection
                    oDistCol.Add dDistanceLineSegToLineSeg2D(oLocationBox.Item("ShtP1"), oLocationBox.Item("ShtP4"), oTextBox.Item("ShtP1"), oTextBox.Item("ShtP2")), "1"
                    oDistCol.Add dDistanceLineSegToLineSeg2D(oLocationBox.Item("ShtP1"), oLocationBox.Item("ShtP4"), oTextBox.Item("ShtP3"), oTextBox.Item("ShtP4")), "2"
                    oDistCol.Add dDistanceLineSegToLineSeg2D(oLocationBox.Item("ShtP2"), oLocationBox.Item("ShtP3"), oTextBox.Item("ShtP1"), oTextBox.Item("ShtP2")), "3"
                    oDistCol.Add dDistanceLineSegToLineSeg2D(oLocationBox.Item("ShtP2"), oLocationBox.Item("ShtP3"), oTextBox.Item("ShtP3"), oTextBox.Item("ShtP4")), "4"
                    For j = 2 To oDistCol.Count
                        If oDistCol.Item(j) < oDistCol.Item(1) Then
                            oDistCol.Add oDistCol.Item(j), , 1
                            oDistCol.Remove (CStr(j))
                        End If
                    Next
                    If oDistCol.Item(1) < dMindist Then
                        iTextNumber = k
                        sLocationColName = oLocationCol.Item(1)
                        dMindist = oDistCol.Item(1)
                    End If
                End If
            Next
            k = k + 1
        Next
        If sLocationColName <> "" Then
            If Not bExistInCol(ViewOrCalloutCol.Item(i), "EXISTINGLOCATIONTXT") Then
                ViewOrCalloutCol.Item(i).Add oListofLocationInfo.Item(sLocationColName), "EXISTINGLOCATIONTXT"
                ViewOrCalloutCol.Item(i).Add iTextNumber, "TEXTNUMBER"
                Set oLocationCol = oListofLocationInfo.Item(sLocationColName)
                oListofLocationInfo.Remove (sLocationColName)
                '***Find Zone location for existing text
                Set oLocationBox = oLocationCol.Item("POINTCOLLECTION").Item(1)
                p1(0) = oLocationBox.Item(1)(0): p1(1) = oLocationBox.Item(1)(1)
                p2(0) = oLocationBox.Item(2)(0)
                p4(1) = oLocationBox.Item(4)(1)
                '***Find center of location info box
                dSheetX = p1(0) + (p2(0) - p1(0)) / 2
                dSheetY = p4(1) + (p1(1) - p4(1)) / 2
                '***Set "ZONE" value to zone of Existing Text
                ViewOrCalloutCol.Item(i).Remove "ZONE"
                ViewOrCalloutCol.Item(i).Add sZoneOfTextInDrawing(oLocationCol.Item("TEXTOBJECT").Item(1), dSheetX, dSheetY), "ZONE"
            End If
        End If
    Next
End Sub
'*** Find text height and width
Public Function dDrawingTextwidth(ByRef oText As DrawingText, Optional ByRef dTextHeight As Double) As Double
    Dim dWidth As Double, dHeight As Double, dLineHeight As Double
    Dim dFontsize As Double
    Dim dWrappingwidth As Double
    Dim dWint As Double
    Dim j As Integer, k As Integer
    Dim dScale  As Double
    Dim lFont As Long
    Dim dFontCorrection As Double

    dScale = oText.Parent.Parent.Scale
    On Error Resume Next
        dWrappingwidth = oText.WrappingWidth
        If Err.Number <> 0 Then dWrappingwidth = 0
    On Error GoTo 0
    If dWrappingwidth <> 0 Then
        dDrawingTextwidth = dWrappingwidth
    End If
    dWidth = 0: dWint = 0: k = 0: dHeight = 0: dLineHeight = 0
    For j = 1 To Len(oText.Text)
       If Mid(oText.Text, j, 1) <> vbLf Then
           lFont = oText.GetParameterOnSubString(catFontName, j, 1)
           If oText.GetParameterOnSubString(catSubscript, j, 1) = 1 Or oText.GetParameterOnSubString(catSuperscript, j, 1) = 1 Then dRatio = 0.605 Else dRatio = 1
           If lFont = 3 Then dFontCorrection = 1.15 Else dFontCorrection = 0.927

           dWint = (oText.GetParameterOnSubString(catFontSize, j, 1) * dFontCorrection * dRatio) / 1000 + dWint
           k = k + 1
           dFontsize = oText.GetParameterOnSubString(catFontSize, j, 1) / 1000
           If dLineHeight < dFontsize Then dLineHeight = dFontsize
       Else
            dHeight = dHeight + dLineHeight
            If dWint > dWidth Then dWidth = dWint
            dWint = 0: k = 0: dLineHeight = 0
       End If
    Next
    If dWint > dWidth Then dWidth = dWint
    dHeight = dHeight + dLineHeight
    
    If dWidth < dWrappingwidth Then
        dDrawingTextwidth = dWidth / dScale
    Else
        If dWrappingwidth <> 0 Then
            dDrawingTextwidth = dWrappingwidth / dScale
        Else
            dDrawingTextwidth = dWidth / dScale
        End If
    End If
    
    dTextHeight = (dHeight + (UBound(Split(oText.Text, vbLf))) * 2.032 + 3.45) / dScale
    
End Function
Public Function oTextBoundingBox(ByRef oText As DrawingText, ByVal dWdth As Double, ByVal dHight As Double) As Collection
    Dim xpos As Double, ypos As Double
    Dim p1(1) As Double, p2(1) As Double, p3(1) As Double, p4(1) As Double
    Dim ShtP1(1) As Double, ShtP2(1) As Double, ShtP3(1) As Double, ShtP4(1) As Double
    Dim oColl As Collection
    Dim oDrawingView As DrawingView
    Dim dViewX As Double
    Dim dViewY As Double
    Dim dViewScale As Double
    Dim dViewAngle As Double

    '*** assumption text not at an angle to sheet axis
    '*   p1------p2
    '*   |       |
    '*   |       |
    '*   p4------p3
    
    Set oDrawingView = oText.Parent.Parent
    
    dViewX = oDrawingView.xAxisData
    dViewY = oDrawingView.yAxisData
    dViewScale = oDrawingView.Scale
    dViewAngle = oDrawingView.Angle
    
    Set oColl = New Collection
    dWidth = dWdth * dViewScale
    dHeight = dHight * dViewScale
    xpos = dViewX + (oText.X * Cos(dViewAngle) - oText.Y * Sin(dViewAngle)) * dViewScale
    ypos = dViewY + (oText.X * Sin(dViewAngle) + oText.Y * Cos(dViewAngle)) * dViewScale
    Select Case oText.AnchorPosition
        Case CatTextAnchorPosition.catTopLeft
            p1(0) = xpos: p1(1) = ypos
            p2(0) = xpos + dWidth: p2(1) = p1(1)
            p3(0) = p2(0): p3(1) = ypos - dHeight
            p4(0) = p1(0): p4(1) = p3(1)
        Case CatTextAnchorPosition.catTopCenter
            p1(0) = xpos - (dWidth / 2): p1(1) = ypos
            p2(0) = xpos + (dWidth / 2): p2(1) = p1(1)
            p3(0) = p2(0): p3(1) = ypos - dHeight
            p4(0) = p1(0): p4(1) = p3(1)
        Case CatTextAnchorPosition.catTopRight
            p1(0) = xpos - dWidth: p1(1) = ypos
            p2(0) = xpos: p2(1) = p1(1)
            p3(0) = p2(0): p3(1) = ypos - dHeight
            p4(0) = p1(0): p4(1) = p3(1)
        Case CatTextAnchorPosition.catBottomLeft
            p1(0) = xpos: p1(1) = ypos + dHeight
            p2(0) = xpos + dWidth: p2(1) = p1(1)
            p3(0) = p2(0): p3(1) = ypos
            p4(0) = p1(0): p4(1) = p3(1)
        Case CatTextAnchorPosition.catBaseLeft
            p1(0) = xpos: p1(1) = ypos + dHeight
            p2(0) = xpos + dWidth: p2(1) = p1(1)
            p3(0) = p2(0): p3(1) = ypos
            p4(0) = p1(0): p4(1) = p3(1)
        Case CatTextAnchorPosition.catBottomCenter
            p1(0) = xpos - (dWidth / 2): p1(1) = ypos + dHeight
            p2(0) = xpos + (dWidth / 2): p2(1) = p1(1)
            p3(0) = p2(0): p3(1) = ypos
            p4(0) = p1(0): p4(1) = p3(1)
        Case CatTextAnchorPosition.catBaseCenter
            p1(0) = xpos - (dWidth / 2): p1(1) = ypos + dHeight
            p2(0) = xpos + (dWidth / 2): p2(1) = p1(1)
            p3(0) = p2(0): p3(1) = ypos
            p4(0) = p1(0): p4(1) = p3(1)
        Case CatTextAnchorPosition.catBottomRight
            p1(0) = xpos - dWidth: p1(1) = ypos + dHeight
            p2(0) = xpos: p2(1) = p1(1)
            p3(0) = p2(0): p3(1) = ypos
            p4(0) = p1(0): p4(1) = p3(1)
        Case CatTextAnchorPosition.catBaseRight
            p1(0) = xpos - dWidth: p1(1) = ypos + dHeight
            p2(0) = xpos: p2(1) = p1(1)
            p3(0) = p2(0): p3(1) = ypos
            p4(0) = p1(0): p4(1) = p3(1)
        Case CatTextAnchorPosition.catMiddleLeft
            p1(0) = xpos: p1(1) = ypos + (dHeight / 2)
            p2(0) = xpos + dWidth: p2(1) = p1(1)
            p3(0) = p2(0): p3(1) = ypos - (dHeight / 2)
            p4(0) = p1(0): p4(1) = p3(1)
        Case CatTextAnchorPosition.catMiddleCenter
            p1(0) = xpos - (dWidth / 2): p1(1) = ypos + (dHeight / 2)
            p2(0) = xpos + (dWidth / 2): p2(1) = p1(1)
            p3(0) = p2(0): p3(1) = ypos - (dHeight / 2)
            p4(0) = p1(0): p4(1) = p3(1)
        Case CatTextAnchorPosition.catMiddleRight
            p1(0) = xpos - dWidth: p1(1) = ypos + (dHeight / 2)
            p2(0) = xpos: p2(1) = p1(1)
            p3(0) = p2(0): p3(1) = ypos - (dHeight / 2)
            p4(0) = p1(0): p4(1) = p3(1)
        Case Else   ' for the sake of completion
            p1(0) = xpos: p1(1) = ypos
            p2(0) = xpos + dWidth: p2(1) = p1(1)
            p3(0) = p2(0): p3(1) = ypos - dHeight
            p4(0) = p1(0): p4(1) = p3(1)
    End Select
    
    '*** Special case for Location info Box
    If bIsLocationInfoText(oText) Then
        Dim xMid As Double, yMid As Double
        Dim dAdjust As Double
        xMid = p1(0) + (p2(0) - p1(0)) / 2
        yMid = p4(1) + (p1(1) - p4(1)) / 2
        
        dAdjust = 8
        p1(0) = xMid - dAdjust: p1(1) = yMid + dAdjust
        p2(0) = xMid + dAdjust: p2(1) = yMid + dAdjust
        p3(0) = xMid + dAdjust: p3(1) = yMid - dAdjust
        p4(0) = xMid - dAdjust: p4(1) = yMid - dAdjust
    
    End If
    
    oColl.Add p1, "ShtP1"
    oColl.Add p2, "ShtP2"
    oColl.Add p3, "ShtP3"
    oColl.Add p4, "ShtP4"
    Set oTextBoundingBox = oColl
End Function
Private Function dDistanceLineSegToLineSeg2D(p1 As Variant, p2 As Variant, p3 As Variant, p4 As Variant) As Double
'********--------------------
'***
'*** returns shortest distance between two line segments Line 1 is from P1 to P2 and Line 2 is from P3 to P4
'*** Line 1 is from P1 to P2 and Line 2 is from P3 to P4
'*** Adapted from http://geomalgorithms.com/a07-_distance.html#dist3D_Segment_to_Segment()
'***
'********--------------------
    Dim U(1) As Double
    Dim V(1) As Double
    Dim W(1) As Double
    Dim Result(1) As Double

    Dim a As Double, b As Double, c As Double, d As Double, DD As Double, sc As Double, sn As Double, sd As Double
    Dim tc As Double, tN As Double, tD As Double
    
    U(0) = p2(0) - p1(0): U(1) = p2(1) - p1(1)
    V(0) = p4(0) - p3(0): V(1) = p4(1) - p3(1)
    W(0) = p1(0) - p3(0): W(1) = p1(1) - p3(1)

    a = DotProd2D(U, U)   ' should be > = zero
    b = DotProd2D(U, V)
    c = DotProd2D(V, V)   ' should be > = zero
    d = DotProd2D(U, W)
    E = DotProd2D(V, W)
    
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
    dDistanceLineSegToLineSeg2D = LengthVector2D(Result) ' Length of Vector

End Function
Public Function DotProd2D(v1 As Variant, v2 As Variant) As Double
'*** dot product of two vectors of 2 elements
    DotProd2D = v1(0) * v2(0) + v1(1) * v2(1)
End Function
Public Function LengthVector2D(v1 As Variant) As Double
'*** Magnitude of vector of 2 elements
    LengthVector2D = Sqr(v1(0) ^ 2 + v1(1) ^ 2)
End Function
Private Sub SendToExcel(ByRef oListOCallouts As Collection, ByRef oListOfViewsInfo As Collection, ByRef oListofLocationInfo As Collection)
    Dim iNbRow As Long
    Dim iNbColumn As Integer
    'Dim iRowCounter As Integer
    Dim ReportCol As Collection
    Dim ReportKeysCol As Collection
    Dim sExportResultToExcel As Variant
    Dim bError As Boolean
    Dim sPropertiesThickness As String
    Dim oWorksheet ' As Worksheet
    Dim oTemplateWorksheet ' As Worksheet
    Dim oWorkbook 'As Workbook
    Dim sdummy As String
    Dim oXcell As clsExcel
    Dim i As Long, j As Integer, k As Integer
    
    '***** Launch Excel
    Set oXcell = New clsExcel
    oXcell.GetNewExcel
    oXcell.ShowExcelWindow
    Set oWorkbook = oXcell.App.Workbooks.Open(sExcelReportTemplate)
    Set oWorksheet = oXcell.App.ActiveSheet
    oWorksheet.Name = "LocationBoxReport"
    
    oXcell.App.ScreenUpdating = False
    
    Dim oCallout As Collection
    Dim oView As Collection
    Dim oLocationTxt As Collection
    For i = 1 To oListOCallouts.Count
        j = i + 1
        oWorksheet.Cells(j, 2).Value = i
        oWorksheet.Cells(j, 1).Value = "ListOfCallout;" & CStr(i)
        Set oCallout = oListOCallouts.Item(i)
        oWorksheet.Cells(j, 3).Value = oCallout.Item("DISPLAYNAME")
        If bExistInCol(oCallout, "EXISTINGLOCATIONTXT") Then
            oWorksheet.Cells(j, 4).Value = oCallout.Item("EXISTINGLOCATIONTXT").Item("ZONE")
        End If
        
        If bExistInCol(oCallout, "ASSOCIATEDVIEW") Then
            Set oView = oCallout.Item("ASSOCIATEDVIEW")
            oWorksheet.Cells(j, 7).Value = oView.Item("DISPLAYNAME")
            If bExistInCol(oView, "EXISTINGLOCATIONTXT") Then
                oWorksheet.Cells(j, 8).Value = oView.Item("EXISTINGLOCATIONTXT").Item("ZONE")
            End If
            '***Accepted location callout values
            oWorksheet.Cells(j, 9).Value = oCallout.Item("ZONE")
            '***value for callout
            oWorksheet.Cells(j, 5).Value = oView.Item("ZONE")
        End If
    Next
    j = oListOCallouts.Count + 1
    
    For i = 1 To oListOfViewsInfo.Count
        j = j + 1
        Set oView = oListOfViewsInfo.Item(i)
        oWorksheet.Cells(j, 1).Value = "ListOfViewsInfo;" & CStr(i)
        oWorksheet.Cells(j, 2).Value = j - 1
        oWorksheet.Cells(j, 7).Value = oView.Item(1)
        If bExistInCol(oView, "EXISTINGLOCATIONTXT") Then
            oWorksheet.Cells(j, 8).Value = oView.Item("EXISTINGLOCATIONTXT").Item("ZONE")
        End If
    Next
    j = oListOCallouts.Count + oListOfViewsInfo.Count + 1

    For i = 1 To oListofLocationInfo.Count
        j = j + 1
        oWorksheet.Cells(j, 1).Value = "ListofLocationInfo;" & CStr(i)
        oWorksheet.Cells(j, 2).Value = j - 1
        oWorksheet.Cells(j, 4).Value = oListofLocationInfo.Item(i).Item("ZONE")
    Next
        
    '---Add status and change color---
    For i = 2 To oListOCallouts.Count + oListOfViewsInfo.Count + oListofLocationInfo.Count + 1
        Call VFNCheckExcelRow(i, oWorksheet, True)
    Next
    '***add freeze pane first row
    'oWorksheet.Activate
    oWorksheet.Rows("2:2").Select
    oXcell.App.ActiveWindow.FreezePanes = True
    '**Fit Column
    oWorksheet.Range("A:J").Select
    oXcell.App.Selection.EntireColumn.AutoFit
'    oWorksheet.Range("A:A").EntireColumn.Hidden = True
'    oXcell.App.cells(1, 2).Select
    '****Hide useless rows of the sheet
    Dim Lastrow As Long
    j = oListOCallouts.Count + oListOfViewsInfo.Count + oListofLocationInfo.Count + 2
    Lastrow = oWorksheet.Cells(oXcell.App.Rows.Count, 1).End(-4121).Row  ' Const xlUp = -4121
    oWorksheet.Range("A" & CStr(j) & ":A" & Lastrow).EntireRow.Hidden = True
    
    '***delete all other sheets in workbook
    For i = 1 To oWorkbook.Sheets.Count
        If oWorkbook.Sheets.Item(i).Name <> oWorksheet.Name Then
           oWorkbook.Sheets.Item(i).Name.Sheets.Delete oWorkbook.Sheets.Item(i).Name
        End If
    Next
    '***Lock sheet & workbook
    oXcell.App.DisplayAlerts = False
    On Error Resume Next
    oWorksheet.Protect sEXCELPWD, True, True
    oWorkbook.Protect sEXCELPWD, True, True
    oWorkbook.bNoRunBeforesave = True
    sReportName = "C:\Temp\LocationAnalysisReport_" & Replace(oDrawingDocument.Name, ".", "_") & Format((DateTime.Timer) / 86400, "hh_mm_ss") & ".xls"
    oWorkbook.SaveAs sReportName
    oWorkbook.bNoRunBeforesave = False
    On Error GoTo 0
    oXcell.App.DisplayAlerts = True
    oXcell.App.ScreenUpdating = True
    'Set oXcell = Nothing
    frmLocationInfo.LaunchDisplayInCatia oXcell, oWorksheet, CATIA.ActiveDocument, oListOCallouts, oListOfViewsInfo, oListofLocationInfo, sSourceWindow
End Sub
Public Function dXcoodforViewTitleLocationTxt(ByVal dXSheetCood As Double, ByVal sTxt As String) As Double
   Dim dLength As Double
   'Assumption font size for view name is always 6.35 and font type SSS1
   dLength = (Len(sTxt) * 6.35 * 0.927) / 2 + 2.54 + 6.35
   dXcoodforViewTitleLocationTxt = dXSheetCood + dLength
End Function
Public Sub VFNCheckExcelRow(ByVal iRow As Long, ByRef oWorksheet, Optional bskip As Boolean = False)
  Dim lerr As Long
  Dim iObject

  
  On Error Resume Next
    Set iObject = oWorksheet.Parent
    lerr = Err.Number
  On Error GoTo 0
  If Not bskip Then Call frmLocationInfo.ModifyExcelSheet(True)
  If lerr <> 0 Then Exit Sub
  If oWorksheet.Cells(iRow, iACCPLCNCALL) = "" Or oWorksheet.Cells(iRow, iEXSTLCNCALL) = "" Then
    oWorksheet.Cells(iRow, iSTATUSCALL) = "NOINFO"
  Else
    If oWorksheet.Cells(iRow, iACCPLCNCALL) = oWorksheet.Cells(iRow, iEXSTLCNCALL) Then
      oWorksheet.Cells(iRow, iSTATUSCALL) = "OK"
    Else
      oWorksheet.Cells(iRow, iSTATUSCALL) = "KO"
    End If
  End If
  
  If oWorksheet.Cells(iRow, iACCPLCNVUE) = "" Or oWorksheet.Cells(iRow, iEXSTLCNVUE) = "" Then
    oWorksheet.Cells(iRow, iSTATUSVUE) = "NOINFO"
  Else
    If oWorksheet.Cells(iRow, iACCPLCNVUE) = oWorksheet.Cells(iRow, iEXSTLCNVUE) Then
      oWorksheet.Cells(iRow, iSTATUSVUE) = "OK"
    Else
      oWorksheet.Cells(iRow, iSTATUSVUE) = "KO"
    End If
  End If
  
  If oWorksheet.Cells(iRow, iSTATUSCALL) = "OK" Then
    oWorksheet.Cells(iRow, iSTATUSCALL).Interior.ColorIndex = 4
  ElseIf oWorksheet.Cells(iRow, iSTATUSCALL) = "KO" Then
    oWorksheet.Cells(iRow, iSTATUSCALL).Interior.ColorIndex = 3
  Else
    oWorksheet.Cells(iRow, iSTATUSCALL).Interior.ColorIndex = 6
  End If
    
  If oWorksheet.Cells(iRow, iSTATUSVUE) = "OK" Then
    oWorksheet.Cells(iRow, iSTATUSVUE).Interior.ColorIndex = 4
  ElseIf oWorksheet.Cells(iRow, iSTATUSVUE) = "KO" Then
    oWorksheet.Cells(iRow, iSTATUSVUE).Interior.ColorIndex = 3
  Else
    oWorksheet.Cells(iRow, iSTATUSVUE).Interior.ColorIndex = 6
  End If
  If Not bskip Then Call frmLocationInfo.ModifyExcelSheet(False)
End Sub

